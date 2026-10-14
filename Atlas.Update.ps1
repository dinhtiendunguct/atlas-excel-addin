param([switch]$CheckOnly, [switch]$Unattended)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root = $PSScriptRoot
$stage = $null
$mutex = New-Object Threading.Mutex($false, 'Local\AtlasExcelUpdater')
$ownsLock = $mutex.WaitOne(0)
if (-not $ownsLock) { $mutex.Dispose(); exit 0 }
function Tell($message) {
    if ($Unattended) { Write-Output $message } else { [void][Windows.Forms.MessageBox]::Show($message, 'Atlas Update') }
}
function HttpsUri($value) {
    $uri = [uri]$value
    if (-not $uri.IsAbsoluteUri -or $uri.Scheme -ne 'https' -or $uri.UserInfo) { throw 'Update URLs must use HTTPS.' }
    return $uri
}
try {
    Start-Transcript -Path (Join-Path $root 'update-last.log') -Force | Out-Null
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
    if (-not $Unattended) {
        $answer = [Windows.Forms.MessageBox]::Show("Atlas $next`n$($feed.notes)`n`nTai va cai dat? Hay luu cong viec va dong tat ca Excel sau khi tai xong.", 'Atlas Update', 'YesNo')
        if ($answer -ne 'Yes') { exit 0 }
    }
    # New releases include the complete installer so the updater itself can
    # evolve and repair startup registration as part of the same operation.
    if ($feed.packageUrl -and $feed.packageSha256) {
        $packagePath = Join-Path $root ('package-' + [guid]::NewGuid().ToString('N') + '.zip')
        $packageDir = Join-Path $root ('stage-' + [guid]::NewGuid().ToString('N'))
        try {
            Invoke-WebRequest -Uri (HttpsUri $feed.packageUrl) -OutFile $packagePath -UseBasicParsing -TimeoutSec 120
            if ((Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash -ne $feed.packageSha256) { throw 'Installer checksum mismatch.' }
            [void][IO.Directory]::CreateDirectory($packageDir)
            $package = [IO.Compression.ZipFile]::OpenRead($packagePath)
            try {
                foreach ($requiredFile in @('Atlas.xlam','Atlas.Update.ps1','Install-Atlas.ps1','update-config.json')) {
                    $entry = $package.GetEntry($requiredFile)
                    if (-not $entry) { throw "Missing package file: $requiredFile" }
                    [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, (Join-Path $packageDir $requiredFile))
                }
            } finally { $package.Dispose() }
            if ((Get-FileHash -LiteralPath (Join-Path $packageDir 'Atlas.xlam') -Algorithm SHA256).Hash -ne $feed.sha256) { throw 'Packaged add-in checksum mismatch.' }
            Tell 'Da tai va kiem tra. Luu va dong tat ca Excel; Atlas se tu cai dat va sua dang ky khoi dong, cho toi da 10 phut.'
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $packageDir 'Install-Atlas.ps1') -WaitForExcel
            if ($LASTEXITCODE -ne 0) { throw 'Installer failed. See update log.' }
            Tell "Da cap nhat Atlas $next. Mo lai Excel de su dung."
            exit 0
        } finally {
            if (Test-Path -LiteralPath $packagePath) { Remove-Item -LiteralPath $packagePath }
            foreach ($requiredFile in @('Atlas.xlam','Atlas.Update.ps1','Install-Atlas.ps1','update-config.json')) {
                $cleanupFile = Join-Path $packageDir $requiredFile
                if (Test-Path -LiteralPath $cleanupFile) { Remove-Item -LiteralPath $cleanupFile }
            }
            if (Test-Path -LiteralPath $packageDir) { Remove-Item -LiteralPath $packageDir }
        }
    }
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
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}
