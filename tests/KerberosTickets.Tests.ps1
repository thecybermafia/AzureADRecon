. (Join-Path $PSScriptRoot 'TestHelpers.ps1')
$node=$script:AuditAst.Find({param($n)$n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Initialize-AADRWindowsKerberosInterop'},$true)
$nativeSource=$node.Find({param($n)$n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value.StartsWith('using System;')},$true).Value
Add-Type -TypeDefinition ($nativeSource + @'
public sealed class FakeLsaApi : AzureADRecon.Native.ILsaApi {
    public string FailAt;
    public bool Closed, Freed;
    public int Calls;
    public int Connect(out System.IntPtr h) { h = FailAt == "Connect" ? System.IntPtr.Zero : new System.IntPtr(123); return FailAt == "Connect" ? unchecked((int)0xC0000022) : 0; }
    public int Lookup(System.IntPtr h, ref AzureADRecon.Native.LsaString name, out uint p) { p=42; return FailAt=="Lookup" ? unchecked((int)0xC0000022) : 0; }
    public int Call(System.IntPtr h,uint p,System.IntPtr request,uint size,out System.IntPtr response,out uint length,out int protocol) {
        Calls++;
        var header=(AzureADRecon.Native.KerbSubmitTicketRequest)System.Runtime.InteropServices.Marshal.PtrToStructure(request,typeof(AzureADRecon.Native.KerbSubmitTicketRequest));
        if(header.MessageType!=21 || header.LogonId.LowPart!=0 || header.LogonId.HighPart!=0 || header.KerbCredOffset!=36 || header.KerbCredSize+36!=size) throw new System.Exception("Bad request layout or target session.");
        response=System.Runtime.InteropServices.Marshal.AllocHGlobal(1);length=1;protocol=FailAt=="Protocol" ? unchecked((int)0xC0000022) : 0;
        return FailAt=="Call" ? unchecked((int)0xC0000022) : 0;
    }
    public int FreeResponse(System.IntPtr p) {Freed=true;System.Runtime.InteropServices.Marshal.FreeHGlobal(p);return 0;}
    public int Close(System.IntPtr p) {Closed=true;return 0;}
    public uint ToWinError(int status) {return 5;}
}
'@)
# Structural fixture only; the real Kerberos package validates full credentials.
[byte[]]$envelope=@(0x76,2,0x30,0)
foreach($stage in @('Success','Connect','Lookup','Call','Protocol')) {
    $api=[FakeLsaApi]::new();$api.FailAt=$stage;$failed=$false
    try{[AzureADRecon.Native.KerberosTicketCache]::Import($envelope,$api)}catch{$failed=$true}
    Assert-True ($failed -eq ($stage -ne 'Success')) "$stage status was not checked."
    Assert-True ($api.Closed -eq ($stage -ne 'Connect') -and $api.Freed -eq ($stage -in @('Success','Call','Protocol'))) "$stage leaked native resources."
}
$failed=$false
try{[AzureADRecon.Native.KerberosTicketCache]::Validate([byte[]]@(0x76,10,0x30,0))}catch{$failed=$true}
Assert-True $failed 'A truncated KRB-CRED envelope passed validation.'
$previousOS=$env:OS;$previousCache=$env:KRB5CCNAME
try {
    $env:OS='OfflineNonWindows'
    $failed=$false
    try{Import-AADRWindowsKerberosTicket -TicketBytes $envelope}catch{$failed=$true}
    Assert-True $failed 'Non-Windows KRB-CRED import was accepted.'
    function Import-AADRWindowsKerberosTicket { [CmdletBinding()]param([byte[]]$TicketBytes) $script:importBytes=$TicketBytes;[AzureADRecon.Native.KerberosTicketCache]::Import($TicketBytes,[FakeLsaApi]::new()) }
    function Get-Module {throw 'An explicit ticket must not use a cached module login.'}
    function Get-Command {[CmdletBinding()]param($Name,$CommandType) $null}
    function Invoke-AADRWindowsSSORequest {
        [CmdletBinding()]param($Endpoint,$Body)
        Assert-True ([string]::IsNullOrEmpty($env:KRB5CCNAME)) 'Windows SSO retained a ccache override.'
        '<Envelope><Assertion xmlns="urn:oasis:names:tc:SAML:1.0:assertion" /></Envelope>'
    }
    function Invoke-RestMethod {[CmdletBinding()]param($Uri,$Method,$Body) @{access_token='test-token'}}
    $env:OS='Windows_NT';$env:KRB5CCNAME='previous-cache'
    $output=@(Get-AADRSeamlessSSOToken -TGT ([Convert]::ToBase64String($envelope)) -TenantID example.test)
    Assert-True ($output.Count -eq 1 -and $output[0] -eq 'test-token') 'SSO status messages contaminated the token.'
    Assert-True ($env:KRB5CCNAME -eq 'previous-cache' -and @($script:importBytes|Where-Object{$_ -ne 0}).Count -eq 0) 'SSO did not restore its environment or clear ticket bytes.'
}finally{$env:OS=$previousOS;$env:KRB5CCNAME=$previousCache}
Write-Host 'Kerberos adapter checks passed (OS calls mocked).'
