param([Parameter(Mandatory=$true)][version]$Version, [Parameter(Mandatory=$true)][uri]$DownloadUrl, [string]$Notes = '', [uri]$PackageUrl, [string]$PackagePath)
$ErrorActionPreference = 'Stop'
if ($DownloadUrl.Scheme -ne 'https') { throw 'Download URL must use HTTPS.' }
$file = Join-Path $PSScriptRoot 'Atlas.xlam'
$manifest = [ordered]@{version=$Version.ToString(); url=$DownloadUrl.AbsoluteUri; sha256=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant(); notes=$Notes}
if ($PackageUrl) {
    if ($PackageUrl.Scheme -ne 'https' -or -not $PackagePath) { throw 'Package URL must use HTTPS and have a local package path.' }
    $manifest.packageUrl = $PackageUrl.AbsoluteUri
    $manifest.packageSha256 = (Get-FileHash -LiteralPath $PackagePath -Algorithm SHA256).Hash.ToLowerInvariant()
}
$manifest | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'latest.json') -Encoding UTF8
Write-Host 'Upload Atlas.xlam and latest.json to your release host. Configure feedUrl with the permanent HTTPS latest.json URL.'
