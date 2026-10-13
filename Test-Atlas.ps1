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
