. (Join-Path $PSScriptRoot 'TestHelpers.ps1')
$script:AADRTokenSession=$null
$script:connects=0; $script:providerCalls=0; $script:requests=0
$script:identity='user-1'; $script:tenant='tenant-1'
function Connect-MgGraph { [CmdletBinding()]param($AccessToken,[switch]$NoWelcome) $script:connects++ }
function Disconnect-MgGraph { [CmdletBinding()]param() }
function Invoke-MgGraphRequest {
    [CmdletBinding()]param($Method,$Uri,$OutputType)
    if($script:rejectPreflight -gt 0 -and $Uri -match '/organization\?') { $script:rejectPreflight--; Write-Error 'Invalid token lifetime.' -ErrorId InvalidAuthenticationToken -ErrorAction Stop }
    if($Uri -match '/me\?') { return [pscustomobject]@{id=$script:identity} }
    [pscustomobject]@{value=@([pscustomobject]@{id=$script:tenant})}
}
$provider={param($request)
    $script:providerCalls++
    $script:reason=$request.Reason
    [pscustomobject]@{Token='replacement-token';ExpiresOn=[DateTimeOffset]::UtcNow.AddHours(1)}
}
$lease=ConvertTo-AADRTokenLease @{access_token='opaque';expires_on='1792857600'}
Assert-True ($lease.ExpiresOnUtc -eq [DateTimeOffset]::FromUnixTimeSeconds(1792857600).UtcDateTime -and $lease.ExpiresOnUtc.Kind -eq [DateTimeKind]::Utc) 'Unix expiry was shifted into local time.'
$lease=ConvertTo-AADRTokenLease @{access_token='opaque';expires_in=3600}
Assert-True ($lease.ExpiresOnUtc -gt [datetime]::UtcNow.AddMinutes(59)) 'OAuth relative lifetime was lost.'
$lease=ConvertTo-AADRTokenLease @{Token='opaque';ExpiresOn=[DateTimeOffset]::UtcNow.AddHours(1)}
Assert-True ($lease.Token -is [securestring] -and $lease.ExpiresOnUtc.Kind -eq 'Utc') 'Opaque token metadata handling failed.'
$failed=$false
try { $null=ConvertTo-AADRTokenLease @{Token='opaque';ExpiresOn='bad'} } catch { $failed=$true }
Assert-True $failed 'Invalid expiry metadata was accepted.'

# Expired input is replaced before it reaches Graph.
Initialize-AADRTokenSession -Token @{Token='expired';ExpiresOn=[DateTimeOffset]::UtcNow.AddHours(-1)} -Provider $provider
Assert-True ($script:providerCalls -eq 1 -and $script:connects -eq 1 -and $script:reason -eq 'Startup') 'Expired startup token was sent to Graph or not replaced.'
Update-AADRTokenSession
Assert-True ($script:providerCalls -eq 1) 'Fresh token was needlessly renewed.'
$before=$script:providerCalls
$script:rejectPreflight=1
Initialize-AADRTokenSession -Provider $provider
Assert-True ($script:providerCalls -eq $before+2 -and $script:AADRTokenSession.Ready) 'A Graph-rejected startup token was not replaced and verified.'
$before=$script:providerCalls
$script:AADRTokenSession.ExpiresOnUtc=[datetime]::UtcNow.AddSeconds(299)
$result=Invoke-AADRReadRequest -Query {'ok'}
Assert-True ($result -eq 'ok' -and $script:providerCalls -eq $before+1 -and $script:reason -eq 'Proactive') 'Renewal window was not checked before a read.'
$before=$script:providerCalls
$denied=$false
try { Invoke-AADRReadRequest -Query { Write-Error Authorization_RequestDenied -ErrorAction Stop } } catch { $denied=$true }
Assert-True ($denied -and $script:providerCalls -eq $before) 'Permission denial triggered token renewal.'

# Revocation/unknown lifetime: replay the read once, discard partial output.
$script:requests=0
$result=@(Invoke-AADRReadRequest -Query {
    $script:requests++
    if($script:requests -eq 1) { 'discarded';Write-Error InvalidAuthenticationToken -ErrorAction Stop }
    'complete'
})
Assert-True ($result.Count -eq 1 -and $result[0] -eq 'complete' -and $script:requests -eq 2 -and $script:reason -eq 'Unauthorized') '401 recovery leaked partial rows or failed to replay once.'
$script:requests=0; $before=$script:providerCalls; $failed=$false
try { Invoke-AADRReadRequest -Query {$script:requests++;Write-Error InvalidAuthenticationToken -ErrorAction Stop} } catch { $failed=$_.FullyQualifiedErrorId -match 'AADRGraphAuthenticationFailed' }
Assert-True ($failed -and $script:requests -eq 2 -and $script:providerCalls -eq $before+1) 'Rejected replacement caused an authentication retry loop.'

# A refresh event must not turn complete inventory into Partial.
$directory=Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $directory
try {
    $script:AADRRun=New-TestRun $directory
    $script:AADRTokenSession.ExpiresOnUtc=[datetime]::UtcNow.AddSeconds(1)
    Invoke-AADRCollectorRun -Name Example -Query { Invoke-AADRReadRequest -Query {[pscustomobject]@{Value=1}} } -OutputDirectory $directory -Formats @()
    Assert-True ($script:AADRRun.Collectors[0].Status -eq 'Success' -and @($script:AADRRun.Events|Where-Object Status -eq 'TokenRenewed').Count -eq 1) 'Successful renewal contaminated completeness.'
    $manifest=Get-Content (Join-Path $directory 'CollectionManifest.json') -Raw
    Assert-True ($manifest -notmatch 'replacement-token|expired') 'Manifest contains token material.'
} finally { Remove-Item $directory -Recurse -Force; $script:AADRRun=$null }

# Fail closed if identity changes, refresh source expires, or stale tokens recur.
$script:identity='other-user';$script:AADRTokenSession.ExpiresOnUtc=[datetime]::UtcNow.AddSeconds(1)
$failed=$false
try { Update-AADRTokenSession } catch { $failed=$_.Exception.Message -match 'different tenant or user' }
Assert-True ($failed -and -not $script:AADRTokenSession.Ready) 'Identity-changing refresh was accepted.'
$script:identity='user-1'
Initialize-AADRTokenSession -Provider $provider
$script:tenant='other-tenant'; $script:AADRTokenSession.ExpiresOnUtc=[datetime]::UtcNow.AddSeconds(1)
$failed=$false
try { Update-AADRTokenSession } catch { $failed=$_.Exception.Message -match 'different tenant or user' }
Assert-True $failed 'Tenant-changing refresh was accepted.'
$script:tenant='tenant-1'
$failed=$false
try { Initialize-AADRTokenSession -Provider { @{Token='stale';expires_in=10} } } catch { $failed=$_.Exception.Message -match 'nearly expired' }
Assert-True $failed 'Provider returned a token inside the refresh window.'
$diagnostic=Get-AADRErrorDiagnostic $Error[0]
Assert-True ($diagnostic.Operation -eq 'TokenMetadata' -and $diagnostic.Reason -eq 'ExpiryWithinRenewalWindow' -and $null -eq $diagnostic.HttpStatus -and $diagnostic.ExpiresOnUtc -and $diagnostic.ClientUtc) 'Local expiry rejection was misreported as a Graph response.'
$failed=$false
try { Initialize-AADRTokenSession -Provider { throw 'secret-provider-message' } } catch { $failed=$_.Exception.Message -notmatch 'secret-provider-message' }
Assert-True $failed 'Provider failure exposed raw sensitive exception text.'
$failed=$false
try { Initialize-AADRTokenSession -Token 'opaque' -ExpiresOn ([datetime]::UtcNow.AddSeconds(-1)) } catch { $failed=$true }
Assert-True $failed 'Expired token without a provider was accepted.'
Initialize-AADRTokenSession -Token 'opaque' -ExpiresOn ([datetime]::UtcNow.AddHours(1)) -WarningAction SilentlyContinue
$script:AADRTokenSession.ExpiresOnUtc=[datetime]::UtcNow.AddSeconds(1)
$failed=$false
try { Invoke-AADRReadRequest -Query {throw 'query should not run'} } catch { $failed=$_.Exception.Message -match 'no AccessTokenProvider' }
Assert-True $failed 'Known expiry without a refresh source reached the query.'
$script:AADRTokenSession=$null

# Driver integration: callback-only login, verified identity, and automatic SSO.
function Get-MgContext { [CmdletBinding()]param() [pscustomobject]@{Account=$null;ClientId=$null;AuthType='UserProvidedAccessToken';Scopes=@()} }
function Get-MgOrganization {
    [CmdletBinding()]param([switch]$All)
    if($script:expireDuringTenant) { $script:AADRTokenSession.ExpiresOnUtc=[datetime]::UtcNow.AddSeconds(1) }
    [pscustomobject]@{Id='tenant-1';DisplayName='Tenant'}
}
function Get-MgDevice { [CmdletBinding()]param([switch]$All) [pscustomobject]@{Id='device';DisplayName='Device'} }
function Get-AADRSeamlessSSOToken {
    [CmdletBinding()]param($TenantID,$Domain,$DC,$TGT,[switch]$IncludeMetadata,[switch]$ForceRefresh)
    $script:ssoCalls++
    $script:ssoForced=$ForceRefresh.IsPresent
    Assert-True ($IncludeMetadata -and $TGT -eq 'same-ticket') 'SSO provider lost its metadata request or ticket parameters.'
    @{Token='sso-token';expires_in=3600}
}
$directory=Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
    Invoke-AzureADRecon -AccessTokenProvider $provider -Collect Tenant -OutputType CSV -AADROutputDir $directory
    $manifest=Get-Content (Join-Path $directory 'CollectionManifest.json') -Raw | ConvertFrom-Json
    Assert-True ($manifest.Status -eq 'Success' -and $manifest.Account -eq 'user-1' -and $null -eq $script:AADRTokenSession) 'Driver failed callback-only login, verified identity attribution, or cleanup.'
} finally { Remove-Item $directory -Recurse -Force -ErrorAction SilentlyContinue }
$directory=Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N'))
try {
    $script:ssoCalls=0;$script:expireDuringTenant=$true
    Invoke-AzureADRecon -SeamlessSSO -TGT same-ticket -Collect Tenant,Devices -OutputType CSV -AADROutputDir $directory
    $manifest=Get-Content (Join-Path $directory 'CollectionManifest.json') -Raw | ConvertFrom-Json
    Assert-True ($manifest.Status -eq 'Success' -and $script:ssoCalls -eq 2 -and $script:ssoForced) 'SSO was not reacquired automatically with the original parameters.'
} finally { Remove-Item $directory -Recurse -Force -ErrorAction SilentlyContinue }

# Cleanup is idempotent after startup failed before a Graph context existed.
$script:cleanupDisconnects=0
$script:AADRCleanupFailed=$false
function Get-MgContext { [CmdletBinding()]param() $null }
function Disconnect-MgGraph { [CmdletBinding()]param() $script:cleanupDisconnects++; throw 'No application to sign out from.' }
Disconnect-AADRGraphSession
Assert-True ($script:cleanupDisconnects -eq 0 -and -not $script:AADRCleanupFailed) 'Cleanup tried to disconnect a nonexistent Graph context.'
function Get-MgContext { [CmdletBinding()]param() [pscustomobject]@{Account='user@example.test'} }
Disconnect-AADRGraphSession
Assert-True ($script:cleanupDisconnects -eq 1 -and -not $script:AADRCleanupFailed) 'Already-disconnected SDK response became a spurious cleanup failure.'
Write-Host 'Token lifecycle checks passed.'
