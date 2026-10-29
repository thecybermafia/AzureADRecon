. (Join-Path $PSScriptRoot 'TestHelpers.ps1')
foreach($parser in @('DomainParser','LicenseParser','ServicePrincipalParser','DirectoryRoleParser','GroupParser','DeviceParser','TenantParser')) {
    $rows=@($script:AADRParserType::$parser([object[]]@([pscustomobject]@{}),1))
    Assert-True ($rows.Count -gt 0) "$parser dropped sparse data."
}
$state=@{}
$rows=@(Invoke-AADRGraphCollection -Name Denied -State $state -WarningAction SilentlyContinue -Query {'partial';Write-Error Authorization_RequestDenied -ErrorAction Stop})
Assert-True ($rows.Count -eq 0 -and -not $state.Available) 'A denied query leaked partial pages.'
$script:sleeps=0; $script:calls=0
function Start-Sleep {param($Seconds) $script:sleeps++}
$result=Invoke-AADRReadRequest -Query {$script:calls++;if($script:calls -lt 3){Write-Error TooManyRequests -ErrorAction Stop};'done'}
Assert-True ($result -eq 'done' -and $script:calls -eq 3 -and $script:sleeps -eq 2) 'Bounded transient retry failed.'
$script:calls=0
$failed=$false
try{$null=Invoke-AADRReadRequest -Query {$script:calls++;Write-Error InvalidAuthenticationToken -ErrorAction Stop}}catch{$failed=$true}
Assert-True ($failed -and $script:calls -eq 1) 'An expired token was retried as a transient error.'
$authException=New-Object System.Exception 'opaque SDK error'
$authException | Add-Member NoteProperty ResponseStatusCode 401
$authRecord=New-Object System.Management.Automation.ErrorRecord $authException,'Unspecified',([System.Management.Automation.ErrorCategory]::NotSpecified),$null
Assert-True (Test-AADRGraphAuthenticationError $authRecord) 'Status-only 401 was not recognized.'
$directory=Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $directory
try {
    $script:AADRRun=New-TestRun $directory
    $caught=$null
    try { Invoke-AADRCollectorRun -Name Users -Query { throw 'Invalid token lifetime.' } -OutputDirectory $directory -Formats @('CSV') } catch { $caught=$_ }
    $manifest=Get-Content (Join-Path $directory 'CollectionManifest.json') -Raw | ConvertFrom-Json
    Assert-True ($caught.FullyQualifiedErrorId -match 'AADRGraphAuthenticationFailed' -and $manifest.Collectors[0].Error -eq 'AuthenticationFailed') 'Token failure was not classified durably in the manifest.'
    Assert-True (@($manifest.Events | Where-Object Status -eq 'AuthenticationFailed').Count -eq 1) 'Authentication failure event was not persisted.'
    $report=[pscustomobject]@{Missing=$null;False=$false;Zero=0;Nested=@{A=@{B=@{C='retained'}}}}
    $csv=Join-Path $directory 'report.csv';$json=Join-Path $directory 'report.json';$xml=Join-Path $directory 'report.xml';$html=Join-Path $directory 'report.html'
    Export-ADRCSV -ADRObj $report -ADFileName $csv
    Export-ADRJSON -ADRObj $report -ADFileName $json
    Export-ADRXML -ADRObj $report -ADFileName $xml
    Export-ADRHTML -ADRObj $report -ADFileName $html
    $fromCsv=Import-Csv $csv;$fromJson=Get-Content $json -Raw|ConvertFrom-Json
    Assert-True ([string]::IsNullOrEmpty($fromCsv.Missing) -and $fromCsv.False -eq 'False' -and $fromCsv.Zero -eq '0') 'CSV changed unknown/false/zero.'
    Assert-True (($fromCsv.Nested | ConvertFrom-Json).A.B.C -eq 'retained') 'CSV lost nested inventory fields.'
    Assert-True ($null -eq $fromJson.Missing -and $fromJson.Nested.A.B.C -eq 'retained') 'JSON truncated nested data.'
    Assert-True ((Test-Path $xml) -and (Test-Path $html)) 'XML/HTML export failed.'
}finally{Remove-Item -LiteralPath $directory -Recurse -Force;$script:AADRRun=$null}
Write-Host 'Collector/export checks passed.'
