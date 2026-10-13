param([switch]$Visible)
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
$qaDir = Join-Path $PSScriptRoot 'qa'
[void][IO.Directory]::CreateDirectory($qaDir)
$preview = Join-Path $qaDir ('Atlas-Demo-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.xlsx')
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open((Join-Path $PSScriptRoot 'Atlas.xlam'))
    $unicodeFolder = Join-Path $qaDir ('Unicode-' + [char]0x1EBF + '-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($unicodeFolder)
    $unicodeUpdater = Join-Path $unicodeFolder 'Atlas.Update.ps1'
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Atlas.Update.ps1') -Destination $unicodeUpdater
    if ($excel.Run("'Atlas.xlam'!UpdaterInFolder", $unicodeFolder) -ne $unicodeUpdater) { throw 'Unicode updater lookup failed' }
    if ([string]$excel.Run("'Atlas.xlam'!UpdaterInFolder", 'https://example.invalid/Atlas') -ne '') { throw 'Cloud URL should not be treated as local path' }
    if ([string]$excel.Run("'Atlas.xlam'!UpdaterInFolder", (Join-Path $qaDir 'missing-folder')) -ne '') { throw 'Missing updater path should be empty' }
    Write-Host '3 updater path regression checks passed'
    $result = $excel.Run("'Atlas.xlam'!AtlasAcceptance", $preview)
    Write-Host $result
    if ($result -like 'FAIL*') { throw $result }
    Write-Host "Demo workbook: $preview"
    $book.Close($false)
    $book = $null
} finally {
    if ($null -ne $excel) {
        $excel.DisplayAlerts = $false
        while ($excel.Workbooks.Count -gt 0) { $excel.Workbooks.Item(1).Close($false) }
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}
