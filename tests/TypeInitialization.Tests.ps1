# Run offline: pwsh -NoProfile -File ./tests/TypeInitialization.Tests.ps1
$ErrorActionPreference = 'Stop'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '../AzureADRecon.ps1'), [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
$sourceAssignment = $ast.Find({ param($n)
    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$MSGraphSource'
}, $true)
Invoke-Expression $sourceAssignment.Extent.Text
foreach ($name in @('Initialize-AADRParserType', 'Initialize-AADRWindowsKerberosInterop', 'Get-AADRDevice', 'Invoke-AADRGraphCollection', 'Test-AADRGraphPermissionError', 'Invoke-AADRReadRequest', 'Update-AADRTokenSession', 'Test-AADRGraphAuthenticationError', 'Stop-AADRGraphAuthentication', 'Test-AADRControlFlowError', 'Get-AADRErrorDiagnostic', 'Test-AADRGraphNotFoundError', 'Get-AADRProperty', 'Add-AADREvent')) {
    $node = $ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name
    }, $true)
    Invoke-Expression $node.Extent.Text
}
function Assert-True($Value, $Message) { if (-not $Value) { throw $Message } }

# Reproduce a session containing the old class before running the fixed initializer.
Add-Type -TypeDefinition 'namespace AADRecon { public static class AzureADClass {} }'
$first = Initialize-AADRParserType -Source $MSGraphSource
$second = Initialize-AADRParserType -Source $MSGraphSource
Assert-True ($first -is [type] -and [object]::ReferenceEquals($first, $second)) 'Identical source was not reused.'
Assert-True ($first.FullName -ne 'AADRecon.AzureADClass') 'An old unversioned class was reused.'

# Changing source must load the new behavior without redefining the old class.
$changedSource = $MSGraphSource.Replace('return ADRObject == null ? 0 : ADRObject.Length;', 'return ADRObject == null ? 0 : ADRObject.Length + 1;')
$changed = Initialize-AADRParserType -Source $changedSource
Assert-True ($changed.FullName -ne $first.FullName) 'Changed source reused a stale parser.'
Assert-True ($first::ObjectCount([object[]]@(1,2)) -eq 2 -and $changed::ObjectCount([object[]]@(1,2)) -eq 3) 'Updated parser behavior did not take effect.'
Assert-True ([object]::ReferenceEquals($changed, (Initialize-AADRParserType -Source $changedSource))) 'Changed version did not become reusable.'
Assert-True ([object]::ReferenceEquals($first, (Initialize-AADRParserType -Source $MSGraphSource))) 'Switching back attempted to redefine the first version.'

# Exercise the same dynamic type binding used by real collectors.
$script:AADRParserType = $first
function Get-MgDevice {
    [CmdletBinding()] param([switch]$All)
    [pscustomobject]@{DisplayName='Test device'; AccountEnabled=$false}
}
$devices = @(Get-AADRDevice -Method MSGraph)
Assert-True ($devices.Count -eq 1 -and $devices[0].DisplayName -eq 'Test device' -and $devices[0].AccountEnabled -eq $false) 'Collector could not use the initialized parser.'
$script:AADRParserType = $changed
$devices = @(Get-AADRDevice -Method MSGraph)
Assert-True ($devices.Count -eq 1 -and $null -eq $devices[0].IsManaged) 'Collector failed with a second parser version.'

# Compilation errors must terminate initialization and leave caller preferences intact.
$originalDebug = $DebugPreference
$rejected = $false
try { $null = Initialize-AADRParserType -Source "namespace AADRecon`n{ invalid source }" } catch { $rejected = $true }
Assert-True ($rejected -and $DebugPreference -eq $originalDebug) 'Compilation failure was hidden or changed preferences.'

# The separate Kerberos interop initializer already guards against repeat loads.
Initialize-AADRWindowsKerberosInterop
$native = 'AzureADRecon.Native.KerberosTicketCache' -as [type]
Initialize-AADRWindowsKerberosInterop
Assert-True ([object]::ReferenceEquals($native, ('AzureADRecon.Native.KerberosTicketCache' -as [type]))) 'Kerberos interop was redefined.'
Write-Host 'Repeated type initialization checks passed.'
