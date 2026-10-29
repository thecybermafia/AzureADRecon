. (Join-Path $PSScriptRoot 'TestHelpers.ps1')
$directory=Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $directory
$null=New-Item -ItemType Directory -Path (Join-Path $directory 'CSV-Files')
$null=New-Item -ItemType Directory -Path (Join-Path $directory 'JSON-Files')
try {
    $script:AADRRun=New-TestRun $directory
    # Nonterminating errors must not produce a successful partial pipeline.
    Invoke-AADRCollectorRun -Name Broken -Query { [pscustomobject]@{Value='discarded'}; Write-Error 'private-token-shaped-value' } -OutputDirectory $directory -Formats CSV
    Invoke-AADRCollectorRun -Name Healthy -Query { [pscustomobject]@{Value='complete'} } -OutputDirectory $directory -Formats CSV
    Assert-True ($script:AADRRun.Collectors[0].Status -eq 'Failed' -and $null -eq $script:AADRRun.Collectors[0].RowCount -and $script:AADRRun.Collectors[1].Status -eq 'Success') 'Collector isolation swallowed a failure or aborted independent work.'
    Assert-True (-not (Test-Path (Join-Path $directory 'CSV-Files/Broken.csv'))) 'Failed collector exported partial pipeline output.'
    try { throw 'private-token-shaped-value' } catch { $diagnostic=Get-AADRErrorDiagnostic $_ }
    Assert-True (($diagnostic|ConvertTo-Json) -notmatch 'private-token-shaped-value') 'String-throw error ID exposed a secret.'
    $authException=New-Object System.Exception 'Invalid token lifetime.'
    $authException|Add-Member NoteProperty ResponseStatusCode 401
    $requestId=[guid]::NewGuid().ToString()
    $authException|Add-Member NoteProperty ResponseHeaders @{'request-id'=$requestId;'Authorization'='private-header'}
    $authRecord=New-Object System.Management.Automation.ErrorRecord $authException,'InvalidAuthenticationToken',([System.Management.Automation.ErrorCategory]::AuthenticationError),$null
    try { Invoke-AADRReadRequest -Query {throw $authRecord} } catch { $diagnostic=Get-AADRErrorDiagnostic $_ }
    Assert-True ($diagnostic.RequestId -eq $requestId -and $diagnostic.HttpStatus -eq 401 -and ($diagnostic|ConvertTo-Json) -notmatch 'private-header') 'Authentication normalization discarded safe correlation details or exposed headers.'

    # Requests can fail independently within the same collector.
    $state=@{}
    $rows=@(Invoke-AADRGraphCollection -Name FailingChild -State $state -Query { 'discarded';throw 'private-message' })
    Assert-True ($rows.Count -eq 0 -and -not $state.Available -and $state.FailureStatus -eq 'Unavailable') 'Request failure became empty/available or leaked partial pages.'

    # A failed checkpoint must not prevent the data from being exported.
    $cacheDir=Join-Path $directory '.checkpoints'
    Remove-Item $cacheDir -Recurse -Force
    Set-Content $cacheDir 'this is a file, not a checkpoint directory'
    Invoke-AADRCollectorRun -Name NoCache -Query { [pscustomobject]@{Value='retained'} } -OutputDirectory $directory -Formats CSV
    $entry=$script:AADRRun.Collectors[-1]
    Assert-True ($entry.Status -eq 'Partial' -and $entry.DataStatus -eq 'Success' -and $null -eq $entry.Checkpoint -and (Test-Path (Join-Path $directory 'CSV-Files/NoCache.csv'))) 'Checkpoint failure lost collected data or created a usable checkpoint claim.'
    Remove-Item $cacheDir -Force

    # Corrupt checkpoints are rejected and recollected, never imported blindly.
    Invoke-AADRCollectorRun -Name Corrupt -Query { [pscustomobject]@{Value='original'} } -OutputDirectory $directory -Formats CSV
    $script:AADRPreviousCollectors=@($script:AADRRun.Collectors[-1])
    Set-Content (Join-Path $cacheDir 'Corrupt.clixml') 'corrupted'
    Invoke-AADRCollectorRun -Name Corrupt -Query { [pscustomobject]@{Value='fresh'} } -OutputDirectory $directory -Formats CSV -Resume
    Assert-True (-not $script:AADRRun.Collectors[-1].Resumed -and (Import-Csv (Join-Path $directory 'CSV-Files/Corrupt.csv')).Value -eq 'fresh') 'Corrupt checkpoint was trusted or stopped recovery.'

    # Failure of one format cannot skip remaining formats.
    $originalExporter=${function:Export-ADR}
    function Export-ADR {
        [CmdletBinding()]param($ADRObj,$AADROutputDir,$OutputType,$ADRModuleName)
        if($OutputType -contains 'CSV') { throw 'private-export-message' }
        Export-ADRJSON -ADRObj $ADRObj -ADFileName (Join-AADRPath @($AADROutputDir,'JSON-Files',($ADRModuleName+'.json')))
    }
    Invoke-AADRCollectorRun -Name Formats -Query {[pscustomobject]@{Value=1}} -OutputDirectory $directory -Formats CSV,JSON
    Assert-True ($script:AADRRun.Collectors[-1].Status -eq 'ExportFailed' -and $script:AADRRun.Collectors[-1].Exports[1].Status -eq 'Success') 'A thrown format failure aborted later formats.'
    Set-Item Function:Export-ADR $originalExporter

    # Atomic output preserves a previous good report when a write fails midway.
    $path=Join-Path $directory 'existing.json'
    Set-Content $path 'good-report'
    $failed=$false
    try { Invoke-AADRAtomicFileWrite -Path $path -Writer {param($target) Set-Content $target 'partial-report';throw 'write failed'} } catch { $failed=$true }
    Assert-True ($failed -and (Get-Content $path -Raw).Trim() -eq 'good-report' -and @(Get-ChildItem $directory -Filter '*.tmp').Count -eq 0) 'Failed atomic write damaged the prior output or leaked temporary files.'
    $mixed=@([pscustomobject]@{Id='unknown';CollectionStatus='Unknown'},[pscustomobject]@{Id='known';DisplayName='Known';Enabled=$false})
    Export-ADRCSV -ADRObj $mixed -ADFileName (Join-Path $directory 'mixed.csv')
    $read=@(Import-Csv (Join-Path $directory 'mixed.csv'))
    Assert-True ($read[1].DisplayName -eq 'Known' -and $read[1].Enabled -eq 'False') 'An unknown first row erased later report columns.'

    # Parser exceptions retain a marker row, including when other workers succeed.
    $rows=@($script:AADRParserType::DeviceParser([object[]]@([pscustomobject]@{DisplayName='valid'},42),2))
    Assert-True ($rows.Count -eq 2 -and @($rows|Where-Object CollectionStatus -eq 'Unknown').Count -eq 1) 'Parser dropped a failed row or aborted other workers.'

    # Reporting and manifest failures must not replace the primary failure.
    $called=$false
    Invoke-AADRReportStep -Name BadReport -Query {throw 'private-report-message'}
    Invoke-AADRReportStep -Name GoodReport -Query {$script:nextReportRan=$true}
    Assert-True $script:nextReportRan 'Report failure prevented a subsequent report operation.'
    $manifestPath=Join-Path $directory 'CollectionManifest.json'
    Remove-Item $manifestPath -Force
    $null=New-Item -ItemType Directory -Path $manifestPath
    Set-Content (Join-Path $manifestPath 'locked') 'prevent replacement by a file'
    $script:AADRManifestFailureRecorded=$false
    Write-AADRManifest
    Invoke-AADRCollectorRun -Name ManifestBroken -Query {throw 'original failure'} -OutputDirectory $directory -Formats @()
    Assert-True ($script:AADRRun.Collectors[-1].Error -eq 'UnexpectedError' -and $script:AADRManifestFailureRecorded) 'Manifest failure replaced the query failure.'
    $failed=$false
    try { Write-AADRManifest -Required } catch { $failed=$_.Exception.Message -match 'initial manifest' }
    Assert-True $failed 'Required manifest failure was ignored.'
    $serialized=$script:AADRRun | ConvertTo-Json -Depth 20
    Assert-True ($serialized -notmatch 'private-token|private-message|private-export|private-report') 'Persisted diagnostics exposed raw exception text.'

    # Cancellation remains fatal, even though ordinary dataset errors continue.
    $cancelled=$false
    try { Invoke-AADRCollectorRun -Name Cancel -Query {throw [OperationCanceledException]::new('cancel')} -OutputDirectory $directory -Formats @() } catch {$cancelled=$true}
    Assert-True $cancelled 'Cancellation was swallowed as a recoverable collector error.'
    Get-ADRExcelComObjRelease -ComObjtoRelease $null -Final $true
    Get-ADRExcelComObjRelease -ComObjtoRelease ([pscustomobject]@{}) -Final $true
} finally { Remove-Item $directory -Recurse -Force -ErrorAction SilentlyContinue; $script:AADRRun=$null }

# Invoke the actual CLI boundary in isolated child processes with offline Graph
# mocks inserted before bootstrap. Exercise exit codes and credential-safe output.
$directory=Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $directory
$originalMode=$env:AADR_TEST_ERROR_MODE
try {
    $source=Get-Content $script:AADRScriptPath -Raw
    $bootstrap=$source.LastIndexOf('$transcribing = $false')
    $mocks=@'
function Connect-MgGraph {
    [CmdletBinding()]param($Scopes,$TenantId,$ClientId,$ContextScope,[switch]$NoWelcome)
    if($env:AADR_TEST_ERROR_MODE -eq 'fatal') { throw 'private-token-shaped-value' }
}
function Disconnect-MgGraph { [CmdletBinding()]param() }
function Get-MgContext { [CmdletBinding()]param() [pscustomobject]@{Account='user@example.test';ClientId='client';AuthType='Delegated'} }
function Invoke-MgGraphRequest { [CmdletBinding()]param($Uri,$Method,$OutputType) [pscustomobject]@{value=@([pscustomobject]@{id='tenant'})} }
function Get-MgOrganization {
    [CmdletBinding()]param([switch]$All)
    if($env:AADR_TEST_ERROR_MODE -eq 'partial') { Write-Error 'private-token-shaped-value'; return }
    [pscustomobject]@{Id='tenant';DisplayName='Offline Tenant'}
}
function Start-Transcript {
    [CmdletBinding()]param($Path)
    if($env:AADR_TEST_ERROR_MODE -eq 'logging') { throw 'private-log-message' }
}
function Stop-Transcript {
    [CmdletBinding()]param()
    if($env:AADR_TEST_ERROR_MODE -in @('cleanup','excelcleanup')) { throw 'private-cleanup-message' }
}
function Export-ADRExcel { [CmdletBinding()]param($ExcelPath,$Logo) $script:AADRExcelFailed=$false }
'@
    $scriptPath=Join-Path $directory 'OfflineCLI.ps1'
    Set-Content $scriptPath ($source.Substring(0,$bootstrap)+$mocks+"`n"+$source.Substring($bootstrap)) -Encoding UTF8
    foreach($mode in @('success','partial','fatal','logging','cleanup','excelcleanup')) {
        $env:AADR_TEST_ERROR_MODE=$mode
        $arguments=@('-NoProfile','-NonInteractive','-File',$scriptPath,'-Collect','Tenant','-OutputType','CSV','-OutputDir',(Join-Path $directory $mode))
        if($mode -in @('logging','cleanup','excelcleanup')) { $arguments+='-Log' }
        if($mode -eq 'excelcleanup') { $arguments+=@('-GenExcel',$directory) }
        $output=(& (Get-Process -Id $PID).Path @arguments 2>&1|Out-String)
        $code=$LASTEXITCODE
        $expected=if($mode -eq 'success'){0}elseif($mode -eq 'fatal'){1}else{2}
        Assert-True ($code -eq $expected -and $output -notmatch 'private-token-shaped-value|private-log|private-cleanup|CategoryInfo|FullyQualifiedErrorId') "CLI $mode exposed a raw exception or returned the wrong exit code: $code. $output"
    }
} finally { $env:AADR_TEST_ERROR_MODE=$originalMode; Remove-Item $directory -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host 'Error isolation, output recovery, and CLI boundary checks passed.'
