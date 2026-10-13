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
        foreach ($moduleFile in @('Atlas.Core.bas', 'Atlas.UI.bas', 'Atlas.Reports.bas', 'Atlas.Tests.bas')) {
            $sourcePath = Join-Path $PSScriptRoot $moduleFile
            $sourceText = [IO.File]::ReadAllText($sourcePath, [Text.Encoding]::UTF8)
            $moduleName = [regex]::Match($sourceText, 'Attribute VB_Name = "([^"]+)"').Groups[1].Value
            $sourceText = [regex]::Replace($sourceText, '(?m)^Attribute VB_Name = .*\r?\n', '')
            # VBA's editor stores ANSI text. Encode Unicode string literals as ASCII
            # hex and decode with U() at runtime, independent of Windows locale.
            $sourceText = [regex]::Replace($sourceText, '"(?:[^"]|"")*"', {
                param($match)
                if ($match.Value -notmatch '[^\x00-\x7F]') { return $match.Value }
                $literal = $match.Value.Substring(1, $match.Value.Length - 2).Replace('""', '"')
                $hex = ($literal.ToCharArray() | ForEach-Object { ([int]$_).ToString('X4') }) -join ''
                return 'U("' + $hex + '")'
            })
            # Split long source lines at argument commas outside string literals.
            $lines = foreach ($line in ($sourceText -split '\r?\n')) {
                $remaining = $line
                while ($remaining.Length -gt 900) {
                    $quoted = $false; $cut = -1
                    for ($i = 0; $i -lt [Math]::Min(850, $remaining.Length); $i++) {
                        if ($remaining[$i] -eq '"') { $quoted = -not $quoted }
                        if (-not $quoted -and $remaining[$i] -eq ',') { $cut = $i + 1 }
                    }
                    if ($cut -lt 0) { throw "VBA line too long in $moduleFile" }
                    $remaining.Substring(0, $cut) + ' _'
                    $remaining = '        ' + $remaining.Substring($cut).TrimStart()
                }
                $remaining
            }
            $module = $components.Add(1)
            $module.Name = $moduleName
            $module.CodeModule.AddFromString(($lines -join "`r`n"))
        }
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
    Add-ZipText 'customUI/customUI.xml' ([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ribbon.xml'), [Text.Encoding]::UTF8))
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
