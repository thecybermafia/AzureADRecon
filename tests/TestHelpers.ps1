$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$script:AADRScriptPath = Join-Path $PSScriptRoot '../AzureADRecon.ps1'
$tokens = $null; $parseErrors = $null
$script:AuditAst = [System.Management.Automation.Language.Parser]::ParseFile($script:AADRScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
$assignment = $script:AuditAst.Find({param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$MSGraphSource'}, $true)
Invoke-Expression $assignment.Extent.Text
foreach ($definition in $script:AuditAst.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] }) {
    Invoke-Expression $definition.Extent.Text
}
$script:AADRParserType = Initialize-AADRParserType -Source $MSGraphSource
$script:AADRRun = $null
function Assert-True($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function New-TestRun([string]$Directory) {
    [pscustomobject]@{OutputDirectory=$Directory; Collectors=(New-Object System.Collections.Generic.List[object]); Events=(New-Object System.Collections.Generic.List[object])}
}
