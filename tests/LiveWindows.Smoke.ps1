# Explicit opt-in only; requires a domain-connected Windows host and real tenant permissions.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$TenantID,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [switch]$SeamlessSSO,
    [switch]$VerifyExcel
)
$ErrorActionPreference='Stop'
if($env:OS -ne 'Windows_NT') { throw 'This smoke test requires Windows.' }
if($env:AADR_LIVE_TESTS -ne '1') { throw 'Set AADR_LIVE_TESTS=1 to explicitly enable live tenant calls.' }
$scriptPath=Join-Path $PSScriptRoot '../AzureADRecon.ps1'
$params=@{TenantID=$TenantID;Collect=@('Tenant','Users','ServicePrincipals');OutputDir=$OutputDirectory;OutputType=@('CSV','JSON')}
if($SeamlessSSO) {
    $params.SeamlessSSO=$true
    if($env:AADR_LIVE_TGT) { $params.TGT=$env:AADR_LIVE_TGT }
}
if($VerifyExcel) { $params.OutputType+= 'EXCEL' }
& $scriptPath @params
$manifest=Get-Content (Join-Path $OutputDirectory 'CollectionManifest.json') -Raw|ConvertFrom-Json
if($manifest.Status -eq 'Failed' -or @($manifest.Collectors|Where-Object Status -eq 'NotStarted').Count) { throw 'Live collection was not completed.' }
if($VerifyExcel -and (@(Get-ChildItem -LiteralPath $OutputDirectory -Filter '*.xlsx').Count -eq 0 -or @($manifest.Events|Where-Object Status -eq 'ExportFailed').Count -gt 0)) { throw 'Excel workbook generation failed.' }
Write-Host "Live run status: $($manifest.Status). Review the manifest for denied or partial datasets."
