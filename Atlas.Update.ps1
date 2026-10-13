param([switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root = $PSScriptRoot
$stage = $null
$mutex = New-Object Threading.Mutex($false, 'Local\AtlasExcelUpdater')
$ownsLock = $mutex.WaitOne(0)
if (-not $ownsLock) { $mutex.Dispose(); exit 0 }
function Tell($message) { [void][Windows.Forms.MessageBox]::Show($message, 'Atlas Update') }
function HttpsUri($value) {
    $uri = [uri]$value
    if (-not $uri.IsAbsoluteUri -or $uri.Scheme -ne 'https' -or $uri.UserInfo) { throw 'Update URLs must use HTTPS.' }
    return $uri
}
try {
    $config = Get-Content -LiteralPath (Join-Path $root 'update-config.json') -Raw | ConvertFrom-Json
    if (-not $config.feedUrl) { Tell 'Chua cau hinh dia chi phat hanh. Hay cung cap link GitHub repository cho Atlas.'; exit 0 }
    # Download bytes: octet-stream assets may otherwise be decoded as ANSI,
    # producing a string with a corrupted UTF-8 BOM on Windows PowerShell.
    $feedPath = Join-Path $root ('feed-' + [guid]::NewGuid().ToString('N') + '.json')
    try {
        Invoke-WebRequest -Uri (HttpsUri $config.feedUrl) -OutFile $feedPath -UseBasicParsing -TimeoutSec 20
        $feed = [IO.File]::ReadAllText($feedPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    } finally {
        if (Test-Path -LiteralPath $feedPath) { Remove-Item -LiteralPath $feedPath }
    }
    if (-not $feed.version) { throw 'Release feed has no version. Cannot determine latest release.' }
    $current = [version]$config.version
    $next = [version]$feed.version
    if ($CheckOnly) {
        [pscustomobject]@{ current = $current.ToString(); latest = $next.ToString(); updateAvailable = ($next -gt $current) } | ConvertTo-Json -Compress | Write-Output
        exit 0
    }
    if ($next -le $current) { Tell "Atlas $current da la phien ban moi nhat."; exit 0 }
    $downloadUri = HttpsUri $feed.url
    if ($feed.sha256 -notmatch '^[a-fA-F0-9]{64}$') { throw 'Release has no valid SHA256.' }
    $answer = [Windows.Forms.MessageBox]::Show("Atlas $next`n$($feed.notes)`n`nTai va cai dat? Hay luu cong viec va dong tat ca Excel sau khi tai xong.", 'Atlas Update', 'YesNo')
    if ($answer -ne 'Yes') { exit 0 }
    $stage = Join-Path $root ('download-' + [guid]::NewGuid().ToString('N') + '.xlam')
    Invoke-WebRequest -Uri $downloadUri -OutFile $stage -UseBasicParsing -TimeoutSec 120
    if ((Get-FileHash -LiteralPath $stage -Algorithm SHA256).Hash -ne $feed.sha256) { throw 'Download checksum mismatch. Installed Atlas was not changed.' }
    $zip = [IO.Compression.ZipFile]::OpenRead($stage)
    try {
        if (-not $zip.GetEntry('xl/vbaProject.bin') -or -not $zip.GetEntry('[Content_Types].xml')) { throw 'Downloaded file is not a VBA add-in.' }
    } finally { $zip.Dispose() }
    Tell 'Da tai va kiem tra. Hay luu cong viec va dong tat ca cua so Excel. Atlas se cho toi da 10 phut; khong tu dong dong Excel.'
    $deadline = (Get-Date).AddMinutes(10)
    while (Get-Process EXCEL -ErrorAction SilentlyContinue) {
        if ((Get-Date) -gt $deadline) { throw 'Excel is still open. Update cancelled; check again later.' }
        Start-Sleep -Seconds 2
    }
    $target = Join-Path $root 'Atlas.xlam'
    $backup = Join-Path $root ('Atlas.previous-' + [guid]::NewGuid().ToString('N') + '.xlam')
    $oldConfig = [IO.File]::ReadAllText((Join-Path $root 'update-config.json'))
    Copy-Item -LiteralPath $target -Destination $backup
    try {
        # Replace on the same volume. Keep the original for manual rollback as well.
        [IO.File]::Replace($stage, $target, $null)
        $stage = $null
        $config.version = $next.ToString()
        $config | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'update-config.json') -Encoding UTF8
    } catch {
        Copy-Item -LiteralPath $backup -Destination $target -Force
        [IO.File]::WriteAllText((Join-Path $root 'update-config.json'), $oldConfig)
        throw
    }
    Tell "Da cap nhat Atlas $next. Ban cu: $backup. Mo lai Excel de su dung."
} catch {
    if ($CheckOnly) { throw }
    Tell "Khong cap nhat duoc: $($_.Exception.Message)"
    exit 1
} finally {
    if ($stage -and (Test-Path -LiteralPath $stage)) { Remove-Item -LiteralPath $stage }
    if ($ownsLock) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
