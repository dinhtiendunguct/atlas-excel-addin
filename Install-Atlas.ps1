param([switch]$WaitForExcel)
$ErrorActionPreference = 'Stop'
if ($WaitForExcel) {
    $deadline = (Get-Date).AddMinutes(10)
    while (Get-Process EXCEL -ErrorAction SilentlyContinue) {
        if ((Get-Date) -gt $deadline) { throw 'Excel is still open. Installation cancelled.' }
        Start-Sleep -Seconds 2
    }
}
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
$officeVersion = '16.0'
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $officeVersion = $excel.Version
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
# AddIns.Installed may not persist the startup entry from an automation instance.
# Change only Atlas entries after Excel has quit; preserve every unrelated add-in.
$optionsKey = "HKCU:\Software\Microsoft\Office\$officeVersion\Excel\Options"
if (-not (Test-Path $optionsKey)) { [void](New-Item -Path $optionsKey -Force) }
$options = Get-ItemProperty -Path $optionsKey
$openEntries = @($options.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' })
$atlasEntries = @($openEntries | Where-Object { [string]$_.Value -match '(?i)(?:^|[\\/])Atlas(?:\[\d+\])?\.xlam"?$' })
$registryBackup = Join-Path $installDir ('startup-backup-' + [guid]::NewGuid().ToString('N') + '.json')
@($openEntries | Select-Object Name,Value) | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $registryBackup -Encoding UTF8
if ($atlasEntries.Count -gt 0) {
    $entryName = $atlasEntries[0].Name
} else {
    $entryName = 'OPEN'; $index = 0
    while ($openEntries.Name -contains $entryName) { $index++; $entryName = 'OPEN' + $index }
}
$startupValue = '"' + (Join-Path $installDir 'Atlas.xlam') + '"'
New-ItemProperty -Path $optionsKey -Name $entryName -Value $startupValue -PropertyType String -Force | Out-Null
foreach ($oldEntry in $atlasEntries) {
    if ($oldEntry.Name -ne $entryName) { Remove-ItemProperty -Path $optionsKey -Name $oldEntry.Name }
}
if ((Get-ItemPropertyValue -Path $optionsKey -Name $entryName) -ne $startupValue) { throw 'Atlas startup registration verification failed.' }
Write-Host "Verified Excel startup: $entryName = $startupValue"
