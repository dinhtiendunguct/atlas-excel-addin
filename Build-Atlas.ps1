$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$excel = $null
$book = $null
$outputPath = Join-Path $PSScriptRoot 'Atlas.xlam'
if (Test-Path -LiteralPath $outputPath) {
    $backupPath = Join-Path $PSScriptRoot ('Atlas.backup-' + [Guid]::NewGuid().ToString('N') + '.xlam')
    Move-Item -LiteralPath $outputPath -Destination $backupPath
    Write-Host "Previous add-in preserved at $backupPath"
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $book = $excel.Workbooks.Add()
    $sourcePath = Join-Path $PSScriptRoot 'Atlas.bas'
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "VBA source file not found: $sourcePath"
    }
    try {
        $project = $book.VBProject
        if ($null -eq $project) { throw 'Excel returned no VBProject.' }
        $components = $project.VBComponents
        if ($null -eq $components) { throw 'Excel returned no VBComponents.' }
        Write-Host "VBA project access OK. Excel $($excel.Version). Importing source..."
    } catch {
        throw "VBA project access failed. Original error: $($_.Exception.ToString())"
    }
    try {
        # VBComponents.Import uses a legacy path conversion that loses Vietnamese characters.
        # Read using .NET Unicode paths and pass source text directly to VBA instead.
        $sourceText = [IO.File]::ReadAllText($sourcePath, [Text.Encoding]::UTF8)
        $sourceText = [regex]::Replace($sourceText, '(?m)^Attribute VB_Name = .*\r?\n', '')
        $module = $components.Add(1)
        $module.Name = 'Atlas'
        $module.CodeModule.AddFromString($sourceText)
        Write-Host 'VBA source loaded successfully.'
    }
    catch { throw "VBA source load failed. Original error: $($_.Exception.ToString())" }
    $book.IsAddin = $true
    $book.SaveAs($outputPath, 55)
    $book.Close($false)
    $book = $null
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
$archive = [IO.Compression.ZipFile]::Open($outputPath, [IO.Compression.ZipArchiveMode]::Update)
function Add-ZipText($name, $content) {
    $entry = $archive.CreateEntry($name)
    $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding($false)))
    try { $writer.Write($content) } finally { $writer.Dispose() }
}
try {
    Add-ZipText 'customUI/customUI.xml' @'
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon><tabs><tab id="AtlasTab" label="Atlas"><group id="AtlasSchedule" label="Tien do"><button id="AtlasNew" label="Tao tien do" size="large" imageMso="TableInsert" onAction="RibbonNew"/><button id="AtlasCalculate" label="Tinh lich va Gantt" size="large" imageMso="CalculateNow" onAction="RibbonCalculate"/><button id="AtlasBaseline" label="Luu baseline" imageMso="FileSave" onAction="RibbonBaseline"/></group><group id="AtlasUpdates" label="Cap nhat"><button id="AtlasUpdate" label="Kiem tra cap nhat" imageMso="RefreshAll" onAction="RibbonUpdate"/></group></tab></tabs></ribbon></customUI>
'@
    $entry = $archive.GetEntry('_rels/.rels')
    $reader = New-Object IO.StreamReader($entry.Open())
    try { [xml]$relationships = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $rel = $relationships.CreateElement('Relationship', $relationships.DocumentElement.NamespaceURI)
    $rel.SetAttribute('Id', 'AtlasRibbon')
    $rel.SetAttribute('Type', 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility')
    $rel.SetAttribute('Target', 'customUI/customUI.xml')
    [void]$relationships.DocumentElement.AppendChild($rel)
    $entry.Delete()
    Add-ZipText '_rels/.rels' $relationships.OuterXml
} finally { $archive.Dispose() }
Write-Host "Created $outputPath"
