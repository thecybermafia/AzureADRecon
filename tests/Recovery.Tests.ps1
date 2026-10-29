. (Join-Path $PSScriptRoot 'TestHelpers.ps1')
$script:tenantCalls=0; $script:deviceCalls=0; $script:disconnectCalls=0; $script:failDevice=$true
function Connect-MgGraph { [CmdletBinding()]param($Scopes,$TenantId,$ClientId,$ContextScope,[switch]$NoWelcome) }
function Disconnect-MgGraph { [CmdletBinding()]param() $script:disconnectCalls++ }
function Get-MgContext { [CmdletBinding()]param() [pscustomobject]@{Account='auditor@example.test';ClientId='client';AuthType='Delegated'} }
function Invoke-MgGraphRequest {
    [CmdletBinding()]param($Uri,$Method,$OutputType)
    [pscustomobject]@{value=@([pscustomobject]@{id='verified-tenant'})}
}
function Get-MgOrganization { [CmdletBinding()]param([switch]$All) $script:tenantCalls++; [pscustomobject]@{Id='verified-tenant'; DisplayName='Tenant'} }
function Get-MgDevice {
    [CmdletBinding()]param([switch]$All)
    $script:deviceCalls++
    if($script:failDevice) { throw 'Simulated unrecoverable device collector error' }
    [pscustomobject]@{Id='device'; DisplayName='Device'; AccountEnabled=$false}
}
$directory=Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
$original=Get-Location
try {
    $failed=$false
    try { Invoke-AzureADRecon -Collect Tenant,Devices -OutputType CSV -AADROutputDir $directory } catch { $failed=$true }
    $manifest=Get-Content (Join-Path $directory 'CollectionManifest.json') -Raw|ConvertFrom-Json
    Assert-True (-not $failed -and $manifest.Status -eq 'Partial' -and $manifest.Collectors[0].Status -eq 'Success' -and $manifest.Collectors[1].Status -eq 'Partial') 'Recoverable failure aborted the run or lost completed/partial collectors.'
    Assert-True ($script:disconnectCalls -eq 1 -and (Get-Location).Path -eq $original.Path) 'Failure bypassed cleanup.'
    $script:failDevice=$false
    Invoke-AzureADRecon -Collect Tenant,Devices -OutputType CSV -AADROutputDir $directory -Resume
    $manifest=Get-Content (Join-Path $directory 'CollectionManifest.json') -Raw|ConvertFrom-Json
    Assert-True ($manifest.Status -eq 'Success' -and $manifest.Collectors[0].Resumed -and $script:tenantCalls -eq 1 -and $script:deviceCalls -eq 2) 'Resume repeated a completed collector or skipped the failed one.'
    Assert-True (Test-Path (Join-AADRPath @($directory,'CSV-Files','Devices.csv'))) 'Cross-platform report path was not created.'
    Assert-True ($script:disconnectCalls -eq 2) 'Successful run bypassed disconnect.'
    # Prevent accidental cross-tenant checkpoint reuse.
    $before=Get-Content (Join-Path $directory 'CollectionManifest.json') -Raw
    $manifest.TenantId='different-tenant'
    $manifest|ConvertTo-Json -Depth 20|Set-Content (Join-Path $directory 'CollectionManifest.json')
    $rejected=$false
    try { Invoke-AzureADRecon -Collect Tenant,Devices -OutputType CSV -AADROutputDir $directory -Resume } catch { $rejected=$true }
    Assert-True $rejected 'A foreign tenant checkpoint was accepted.'
    $preserved=Get-Content (Join-Path $directory 'CollectionManifest.json') -Raw|ConvertFrom-Json
    Assert-True ($preserved.TenantId -eq 'different-tenant' -and $preserved.Collectors.Count -eq 2) 'Rejected resume overwrote the prior manifest.'
    # Empty checkpoints must remain empty when imported.
    $script:AADRRun=New-TestRun $directory
    $script:AADRPreviousCollectors=@()
    Invoke-AADRCollectorRun -Name Empty -Query {} -OutputDirectory $directory -Formats @('CSV')
    $script:AADRPreviousCollectors=$script:AADRRun.Collectors.ToArray()
    $script:AADRRun=New-TestRun $directory
    Invoke-AADRCollectorRun -Name Empty -Query { throw 'Empty checkpoint should have resumed' } -OutputDirectory $directory -Formats @('CSV') -Resume
    Assert-True ($script:AADRRun.Collectors[0].Status -eq 'Empty' -and $script:AADRRun.Collectors[0].RowCount -eq 0) "Empty checkpoint resumed incorrectly: $($script:AADRRun.Collectors[0].Status), rows=$($script:AADRRun.Collectors[0].RowCount)."
    # Partial collectors cannot become completed checkpoints; export failures are persisted.
    $script:AADRRun=New-TestRun $directory
    Invoke-AADRCollectorRun -Name Partial -Query { Add-AADREvent -Operation Test -Status PermissionDenied -Detail 'Denied subquery'; [pscustomobject]@{Value=1} } -OutputDirectory $directory -Formats @('CSV')
    Assert-True ($script:AADRRun.Collectors[0].Status -eq 'Partial' -and -not $script:AADRRun.Collectors[0].Checkpoint) 'Partial result became resumable as complete.'
    $script:AADRRun=New-TestRun $directory
    Export-ADRCSV -ADRObj ([pscustomobject]@{Value=1}) -ADFileName (Join-AADRPath @($directory,'missing','test.csv')) -WarningAction SilentlyContinue
    Assert-True ($script:AADRRun.Events[0].Status -eq 'ExportFailed') 'Export failure was not persisted.'
    # Export failure must not force completed data to be queried again.
    $script:AADRRun=New-TestRun $directory
    $script:AADRPreviousCollectors=@()
    Invoke-AADRCollectorRun -Name ReExport -Query {[pscustomobject]@{Value=1}} -OutputDirectory $directory -Formats @('JSON') -WarningAction SilentlyContinue
    Assert-True ($script:AADRRun.Collectors[0].Status -eq 'ExportFailed' -and $script:AADRRun.Collectors[0].DataStatus -eq 'Success') 'Export status obscured successful collection.'
    $script:AADRPreviousCollectors=$script:AADRRun.Collectors.ToArray()
    $null=New-Item -ItemType Directory -Path (Join-Path $directory 'JSON-Files') -Force
    $script:AADRRun=New-TestRun $directory
    Invoke-AADRCollectorRun -Name ReExport -Query {throw 'Completed data should not be fetched again'} -OutputDirectory $directory -Formats @('JSON') -Resume
    Assert-True ($script:AADRRun.Collectors[0].Resumed -and $script:AADRRun.Collectors[0].Status -eq 'Success') 'Failed export could not resume from complete data.'

} finally { Remove-Item -LiteralPath $directory -Recurse -Force -ErrorAction SilentlyContinue; $script:AADRRun=$null }
Write-Host 'Recovery/manifest checks passed.'
