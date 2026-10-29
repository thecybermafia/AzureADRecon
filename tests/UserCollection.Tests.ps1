. (Join-Path $PSScriptRoot 'TestHelpers.ps1')
$script:scenario='Success'
$script:userCalls=0
$script:methodCalls=0
function Get-MgUser {
    [CmdletBinding()]param([switch]$All,$Property,$PageSize)
    $script:userCalls++
    if($script:scenario -eq 'DeniedSignIn' -and $Property -contains 'signInActivity') { Write-Error Authorization_RequestDenied -ErrorAction Stop }
    $activity=if($Property -contains 'signInActivity') { [pscustomobject]@{lastSignInDateTime=[datetime]'2026-10-23'; lastNonInteractiveSignInDateTime=[datetime]'2026-10-22'; lastSuccessfulSignInDateTime=[datetime]'2026-01-01'} } else {$null}
    [pscustomobject]@{Id='user'; DisplayName='User'; UserPrincipalName='user@example.test'; AccountEnabled=$false; assignedLicenses=@(); createdDateTime=$null; lastPasswordChangeDateTime=$null; signInActivity=$activity}
}
function Get-AADRMFARegistration {
    if($script:scenario -ne 'MissingRegistration') { [pscustomobject]@{id='user'; isMfaRegistered=$true; isMfaCapable=$false} }
}
function Get-MgUserAuthenticationMethod {
    [CmdletBinding()]param($UserId)
    $script:methodCalls++
    if($script:scenario -eq 'ExpiredMethods') { Write-Error 'Invalid token lifetime.' -ErrorId InvalidAuthenticationToken -ErrorAction Stop }
    if($script:scenario -eq 'ThrottledMethods' -and $script:methodCalls -eq 1) { Write-Error TooManyRequests -ErrorAction Stop }
    if($script:scenario -eq 'DeniedMethods') { Write-Error Authentication_RequestFromUnsupportedUserRole -ErrorAction Stop }
    [pscustomobject]@{AdditionalProperties=@{'@odata.type'='#microsoft.graph.phoneAuthenticationMethod';phoneType='mobile';phoneNumber='masked'}}
}
function Collect-Users { @(Get-AADRUser -Method MSGraph -date ([datetime]'2026-10-24') -DormantTimeSpan 30 -PassMaxAge 90 -PageSize 200 -WarningAction SilentlyContinue) }
$user=(Collect-Users)[0]
Assert-True ($user.'Dormant (> 30 days)' -eq $true -and $user.'Logon Age (days)' -eq 1) 'A recent attempt incorrectly made an old successful sign-in active.'
Assert-True ($user.'MFA Registered' -eq $true -and $user.'MFA Capable' -eq $false -and $user.MFAStatus -eq 'Registered' -and $user.'MFA Enforcement' -eq 'NotEvaluated') 'MFA registration was confused with capability or enforcement.'
Assert-True ($null -eq $user.whenCreated -and $null -eq $user.'Password Age (days)' -and $user.Enabled -eq $false -and $user.'Assigned Licenses (Count)' -eq 0) 'Null/false/zero handling regressed.'
$script:scenario='DeniedMethods'
$user=(Collect-Users)[0]
Assert-True ($user.MFAStatus -eq 'Registered' -and $null -eq $user.AuthenticationMethods -and $null -eq $user.'MFA Phone') 'Denied method details invalidated independently available registration facts.'
$script:scenario='DeniedSignIn'
$script:userCalls=0
$user=(Collect-Users)[0]
Assert-True ($script:userCalls -eq 2 -and $null -eq $user.'Dormant (> 30 days)' -and $null -eq $user.LastSuccessfulSignInDateTime) 'Denied sign-in access prevented fallback or fabricated dormancy.'
$script:scenario='MissingRegistration'
$user=(Collect-Users)[0]
Assert-True ($user.MFAStatus -eq 'Unknown' -and $null -eq $user.'MFA Registered') 'Missing registration became not registered.'
$offsetDate=ConvertTo-AADRDateTime ([DateTimeOffset]'2026-10-24T12:00:00+05:30')
Assert-True ($offsetDate -eq ([datetime]'2026-10-24T06:30:00Z').ToUniversalTime() -and $offsetDate.Kind -eq [DateTimeKind]::Utc) 'DateTimeOffset was not normalized to UTC.'
Assert-True ($null -eq (ConvertTo-AADRDateTime 'invalid-date')) 'Invalid date did not remain unknown.'
$script:scenario='ExpiredMethods'; $script:methodCalls=0
$caught=$null
try { $null=Collect-Users } catch { $caught=$_ }
Assert-True ($null -ne $caught -and $caught.FullyQualifiedErrorId -match 'AADRGraphAuthenticationFailed' -and $script:methodCalls -eq 1) 'Token lifetime failure was swallowed, misclassified, or retried.'
Assert-True ($caught.Exception.Message -match 'fresh Microsoft Graph token' -and $caught.Exception.Message -match 'failed collector restarts') 'Token failure omitted actionable recovery limits.'
$script:scenario='ThrottledMethods'; $script:methodCalls=0
function Start-Sleep { param($Seconds) }
$user=(Collect-Users)[0]
Assert-True ($script:methodCalls -eq 2 -and $null -ne $user.AuthenticationMethods) 'A transient method request did not recover.'

# A persistent method-service outage must preserve basic users without retrying
# every user indefinitely. Five failed lookups each get the bounded four attempts.
function Get-MgUser {
    [CmdletBinding()]param([switch]$All,$Property,$PageSize)
    foreach($i in 1..10) { [pscustomobject]@{Id="user-$i";DisplayName="User $i";AccountEnabled=$false;assignedLicenses=@()} }
}
function Get-MgUserAuthenticationMethod {
    [CmdletBinding()]param($UserId)
    $script:methodCalls++
    Write-Error ServiceUnavailable -ErrorAction Stop
}
$script:methodCalls=0
$users=@(Collect-Users)
Assert-True ($users.Count -eq 10 -and $script:methodCalls -eq 20 -and @($users|Where-Object {$_.'Authentication Methods Collection Status' -eq 'Unknown'}).Count -eq 10) 'Method outage aborted users, fabricated empty details, or bypassed the circuit breaker.'

# A malformed property affects one row, not the rest of the user inventory.
Add-Type -TypeDefinition 'public class BadUserField { public override string ToString() { throw new System.FormatException("private-user-field"); } }'
function Get-MgUser {
    [CmdletBinding()]param([switch]$All,$Property,$PageSize)
    $bad=[pscustomobject]@{Id='bad';DisplayName=[BadUserField]::new();accountEnabled=$false}
    $bad
    [pscustomobject]@{Id='good';DisplayName='Good';AccountEnabled=$false;assignedLicenses=@()}
}
function Get-MgUserAuthenticationMethod { [CmdletBinding()]param($UserId) }
$users=@(Collect-Users)
Assert-True ($users.Count -eq 2 -and $users[0].Id -eq 'bad' -and $users[0].CollectionStatus -eq 'Unknown' -and $users[1].DisplayName -eq 'Good') 'Malformed user record was dropped or aborted a healthy user.'
Write-Host 'User collection checks passed.'
