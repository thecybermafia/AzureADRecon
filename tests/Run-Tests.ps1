$ErrorActionPreference='Stop'
foreach($test in Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' | Sort-Object Name) {
    Write-Host "Running $($test.Name)"
    # Isolate function mocks, loaded types, and script-scoped state between suites.
    & (Get-Process -Id $PID).Path -NoProfile -NonInteractive -File $test.FullName
    if($LASTEXITCODE -ne 0) { throw "$($test.Name) failed with exit code $LASTEXITCODE" }
}
