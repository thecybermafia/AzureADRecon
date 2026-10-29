. (Join-Path $PSScriptRoot 'TestHelpers.ps1')
$script:paths=@(); $script:mode='Normal'
function Invoke-MgGraphRequest {
    [CmdletBinding()]param($Uri,$Method,$OutputType)
    Assert-True ($Method -eq 'GET') 'Inventory attempted a mutation.'
    $script:paths += $Uri
    $path=([uri]$Uri).AbsolutePath
    switch -Regex ($path) {
        '/owners$' { return [pscustomobject]@{value=@([pscustomobject]@{id='owner';displayName=$null;userPrincipalName=$null;'@odata.type'='#microsoft.graph.user'})} }
        '/appRoleAssignments$' { return [pscustomobject]@{value=@([pscustomobject]@{id='grant';resourceId='api';resourceDisplayName='API';appRoleId='permission-id'})} }
        '/applications$' {
            $next=if($script:mode -eq 'ForeignNextLink') {'https://untrusted.example/v1.0/applications'}else{$null}
            return [pscustomobject]@{value=@([pscustomobject]@{id='app';appId='client';displayName='Application';requiredResourceAccess=@();passwordCredentials=@([pscustomobject]@{keyId='secret-id';displayName='Secret metadata';endDateTime='2025-01-01T00:00:00Z'});keyCredentials=@()});'@odata.nextLink'=$next}
        }
        '/servicePrincipals$' { return [pscustomobject]@{value=@([pscustomobject]@{id='sp';appId='client';displayName='Enterprise app';passwordCredentials=@();keyCredentials=@()})} }
        '/oauth2PermissionGrants$' { return [pscustomobject]@{value=@([pscustomobject]@{id='consent';clientId='sp';resourceId='api';consentType='AllPrincipals';scope='Read.Test'})} }
        '/subscribedSkus$' { return [pscustomobject]@{value=@([pscustomobject]@{id='sku';consumedUnits=0;prepaidUnits=[pscustomobject]@{enabled=10;warning=0;suspended=$null};servicePlans=@([pscustomobject]@{servicePlanId='plan';servicePlanName='Plan';provisioningStatus='Success'})})} }
        '/userRegistrationDetails$' { return [pscustomobject]@{value=@([pscustomobject]@{id='user';isMfaRegistered=$true;isMfaCapable=$false})} }
        '/roleEligibilityScheduleInstances$' { return [pscustomobject]@{value=@([pscustomobject]@{id='eligible';principalId='user';roleDefinitionId='role';endDateTime=$null})} }
        default { throw "Unexpected test endpoint: $path" }
    }
}
function Get-MgServicePrincipal {
    [CmdletBinding()]param([switch]$All,$Property)
    [pscustomobject]@{id='sp';displayName='Service';oauth2PermissionScopes=@([pscustomobject]@{id='scope';value='Read.Exposed'})}
}
$app=@(Get-AADRApplicationInventory)[0]
Assert-True ($app.Id -eq 'app' -and $app.RequestedResourceAccess -eq '[]') 'Application inventory lost requested permissions.'
$owners=@(Get-AADRObjectOwners -Kind applications)
Assert-True ($owners.Count -eq 1 -and $owners[0].OwnerId -eq 'owner' -and $null -eq $owners[0].OwnerName) 'Limited owner information lost its ID.'
$expiry=@(Get-AADRCredentialExpiry -date ([datetime]'2026-01-01'))
Assert-True ($expiry.Count -eq 1 -and $expiry[0].Expired -eq $true -and $expiry[0].DaysToExpiry -eq -365) 'Credential expiry calculation was incorrect.'
Assert-True (-not ($expiry|ConvertTo-Json -Depth 10).Contains('secretText')) 'Credential inventory included secret material.'
$grants=@(Get-AADRConsentGrants)
Assert-True ($grants[0].scope -eq 'Read.Test' -and $grants[0].clientId -eq 'sp') 'Consent grant collection failed.'
$permissions=@(Get-AADRApplicationPermissions)
Assert-True ($permissions[0].AppRoleId -eq 'permission-id' -and $permissions[0].ClientId -eq 'sp') 'Granted app permissions were not collected.'
$license=@(Get-AADRLicense -Method MSGraph)[0]
Assert-True ($license.ConsumedUnits -eq 0 -and $license.'PrepaidUnits-Enabled' -eq 10 -and $license.'ServicePlans (Count)' -eq 1 -and $null -eq $license.'PrepaidUnits-Suspended') 'Detailed license inventory lost null or known values.'
$sp=@(Get-AADRServicePrincipal -Method MSGraph -date (Get-Date))[0]
Assert-True ($sp.'OAuth2Permissions (Count)' -eq 1 -and $sp.'Exposed OAuth2 Scopes (Count)' -eq 1) 'Exposed scope count remained hard-coded zero.'
$sparse=@($script:AADRParserType::ServicePrincipalParser([object[]]@([pscustomobject]@{Id='sparse'}),1))[0]
Assert-True ($null -eq $sparse.'OAuth2Permissions (Count)') 'Unavailable exposed scopes became zero.'
$pim=@(Invoke-AADRGraphCollection -Name PIMEligibility -Query {Get-AADRGraphPages 'roleManagement/directory/roleEligibilityScheduleInstances'})
Assert-True ($pim[0].principalId -eq 'user' -and $null -eq $pim[0].endDateTime) 'PIM eligibility data was lost.'
$script:mode='ForeignNextLink'
$failed=$false
try {$null=Get-AADRGraphPages 'applications'}catch{$failed=$true}
Assert-True ($failed -and -not ($script:paths -like '*untrusted.example*')) 'A pagination link could send the token to another host.'
Write-Host 'Inventory checks passed.'
