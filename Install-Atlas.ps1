$ErrorActionPreference = 'Stop'
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Save your work and close Excel before installing.' }
$installDir = Join-Path $env:LOCALAPPDATA 'AtlasExcel'
[void][IO.Directory]::CreateDirectory($installDir)
foreach ($name in @('Atlas.xlam', 'Atlas.Update.ps1', 'update-config.json')) {
    $source = Join-Path $PSScriptRoot $name
    if (-not (Test-Path -LiteralPath $source)) { throw "Missing $source. Build Atlas first." }
}
if (Test-Path -LiteralPath (Join-Path $installDir 'Atlas.xlam')) { throw 'Atlas is already installed. Use its updater instead.' }
foreach ($name in @('Atlas.xlam', 'Atlas.Update.ps1', 'update-config.json')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $installDir $name)
}
Write-Host "Installed files to $installDir"
Write-Host "In Excel Add-ins Browse, select: $(Join-Path $installDir 'Atlas.xlam')"
