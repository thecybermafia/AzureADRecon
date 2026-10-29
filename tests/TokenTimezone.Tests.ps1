. (Join-Path $PSScriptRoot 'TestHelpers.ps1')
$definition=$script:AuditAst.Find({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertTo-AADRTokenLease'},$true).Extent.Text
$propertyDefinition=$script:AuditAst.Find({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-AADRProperty'},$true).Extent.Text
$normalizeDefinition=$script:AuditAst.Find({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertTo-AADRGraphAccessToken'},$true).Extent.Text
$directory=Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $directory
$originalTZ=$env:TZ
try {
    $testPath=Join-Path $directory 'TimezoneCheck.ps1'
    $checks=@'
$ErrorActionPreference='Stop'
[TimeZoneInfo]::ClearCachedData()
$expiry=[DateTimeOffset]::UtcNow.AddHours(1).ToUnixTimeSeconds()
$lease=ConvertTo-AADRTokenLease @{access_token='opaque';expires_on=[string]$expiry}
$expected=[DateTimeOffset]::FromUnixTimeSeconds($expiry).UtcDateTime
if($lease.ExpiresOnUtc.Kind -ne [DateTimeKind]::Utc -or $lease.ExpiresOnUtc -ne $expected -or $lease.ExpiresOnUtc -le [DateTime]::UtcNow.AddMinutes(55)) { throw 'Fresh token was shifted by the local UTC offset.' }
Write-Host ('UTC lease conversion passed in ' + [TimeZoneInfo]::Local.Id)
'@
    Set-Content $testPath ($propertyDefinition+"`n"+$normalizeDefinition+"`n"+$definition+"`n"+$checks) -Encoding UTF8
    foreach($zone in @('UTC','America/Los_Angeles','Asia/Kolkata')) {
        # Unix .NET honors TZ; on Windows this still checks the actual machine's
        # zone without modifying system settings or requiring admin rights.
        $env:TZ=$zone
        & (Get-Process -Id $PID).Path -NoProfile -NonInteractive -File $testPath
        Assert-True ($LASTEXITCODE -eq 0) "UTC lease regression failed for TZ=$zone."
    }
} finally { $env:TZ=$originalTZ; Remove-Item $directory -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host 'Token timezone regression checks passed.'
