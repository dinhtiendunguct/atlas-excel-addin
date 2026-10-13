$ErrorActionPreference = 'Stop'
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Save your work and close Excel before installing.' }
$installDir = Join-Path $env:LOCALAPPDATA 'AtlasExcel'
[void][IO.Directory]::CreateDirectory($installDir)
foreach ($name in @('Atlas.xlam', 'Atlas.Update.ps1', 'update-config.json')) {
    $source = Join-Path $PSScriptRoot $name
    if (-not (Test-Path -LiteralPath $source)) { throw "Missing $source. Build Atlas first." }
}
if (Test-Path -LiteralPath (Join-Path $installDir 'Atlas.xlam')) {
    $backup = Join-Path $installDir ('Atlas.previous-' + [guid]::NewGuid().ToString('N') + '.xlam')
    Copy-Item -LiteralPath (Join-Path $installDir 'Atlas.xlam') -Destination $backup
}
foreach ($name in @('Atlas.xlam', 'Atlas.Update.ps1', 'update-config.json')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $installDir $name)
    Unblock-File -LiteralPath (Join-Path $installDir $name)
}
Write-Host "Installed files to $installDir"
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $blank = $excel.Workbooks.Add()
    foreach ($existing in $excel.AddIns) {
        if ($existing.Name -eq 'Atlas.xlam' -and $existing.FullName -ne (Join-Path $installDir 'Atlas.xlam')) {
            $existing.Installed = $false
        }
    }
    $addin = $excel.AddIns.Add((Join-Path $installDir 'Atlas.xlam'), $false)
    $addin.Installed = $true
    $blank.Close($false)
    Write-Host 'Atlas registered. Open Excel and select the Atlas tab.'
} catch {
    Write-Warning "Excel registration failed: $($_.Exception.Message)"
    Write-Host "Fallback: Excel Add-ins Browse -> $(Join-Path $installDir 'Atlas.xlam')"
    exit 1
} finally {
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
