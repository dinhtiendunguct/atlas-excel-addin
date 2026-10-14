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
$officeVersion = '16.0'
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
# Remove obsolete Atlas browse entries so they cannot be selected accidentally.
$managerKey = "HKCU:\Software\Microsoft\Office\$officeVersion\Excel\Add-in Manager"
if (Test-Path $managerKey) {
    $managerEntries = @((Get-ItemProperty $managerKey).PSObject.Properties | Where-Object { $_.Name -match '(?i)(?:^|[\\/])Atlas(?:\[\d+\])?\.xlam$' -and $_.Name -ne (Join-Path $installDir 'Atlas.xlam') })
    if ($managerEntries.Count -gt 0) {
        $managerEntries | Select-Object Name,Value | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $installDir ('manager-backup-' + [guid]::NewGuid().ToString('N') + '.json')) -Encoding UTF8
        foreach ($managerEntry in $managerEntries) { Remove-ItemProperty -Path $managerKey -Name $managerEntry.Name }
    }
}
