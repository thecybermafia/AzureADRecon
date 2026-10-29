. (Join-Path $PSScriptRoot 'TestHelpers.ps1')
$common = @{TenantID='tenant'; ClientID='00000000-0000-0000-0000-000000000001'}
function Get-Item { [CmdletBinding()] param($LiteralPath) [pscustomobject]@{Subject='Mock'} }
function Connect-MgGraph {
    [CmdletBinding(DefaultParameterSetName='User')]
    param(
        [Parameter(ParameterSetName='User')][string[]]$Scopes,
        [Parameter(ParameterSetName='Certificate')][string]$CertificateThumbprint,
        [Parameter(ParameterSetName='Certificate')][string]$CertificateSubjectName,
        [Parameter(ParameterSetName='Certificate')][object]$Certificate,
        [Parameter(ParameterSetName='Secret')][pscredential]$ClientSecretCredential,
        [string]$ClientId, [string]$TenantId, [string]$ContextScope, [switch]$NoWelcome
    )
    $script:boundSet = $PSCmdlet.ParameterSetName
}
foreach ($field in @('CertificateThumbprint','CertificateName','Certificate')) {
    $config = $common.Clone(); $config[$field]='certificate'
    $params = Get-AADRGraphConnectionParameters -Config $config -Scopes @('Directory.Read.All')
    Assert-True (-not $params.ContainsKey('Scopes')) 'Certificate authentication included delegated scopes.'
    Connect-MgGraph @params
    Assert-True ($script:boundSet -eq 'Certificate') 'Certificate parameter set did not bind.'
}
$credential = [pscredential]::new($common.ClientID,(ConvertTo-SecureString 'test-secret' -AsPlainText -Force))
$params=Get-AADRGraphConnectionParameters -Config $common -Credential $credential -Scopes @('Directory.Read.All')
Connect-MgGraph @params
Assert-True ($script:boundSet -eq 'Secret' -and $params.ClientSecretCredential -eq $credential) 'Credential did not authenticate as an application credential.'
$rejected=$false
try { $null=Get-AADRGraphConnectionParameters -Config $common -Credential ([pscredential]::new('user@example.test',$credential.Password)) } catch { $rejected=$true }
Assert-True $rejected 'User/password credential was silently ignored.'
$params=Get-AADRGraphConnectionParameters -Config @{} -Scopes @('User.Read.All')
Connect-MgGraph @params
Assert-True ($script:boundSet -eq 'User') 'Interactive authentication did not bind.'
function Get-MgContext { [CmdletBinding()]param() [pscustomobject]@{Account='actual@example.test'; ClientId='actual-app'; AppName=''; AuthType='Delegated'} }
$about=@(Get-AADRAbout -Method MSGraph -date (Get-Date) -AzureADReconVersion Test -Credential $credential -RanonComputer Test -TotalTime 0)
Assert-True (($about|Where-Object Category -eq 'Ran as user').Value -eq 'actual@example.test') 'Report attribution trusted a supplied credential over the Graph context.'
foreach($token in @('opaque-token',[pscustomobject]@{Token='opaque-token'},(ConvertTo-SecureString 'opaque-token' -AsPlainText -Force))) {
    $actual=ConvertTo-AADRGraphAccessToken -Token $token
    Assert-True ([System.Net.NetworkCredential]::new('',$actual).Password -eq 'opaque-token') 'Token normalization broke.'
}
Write-Host 'Authentication checks passed.'
