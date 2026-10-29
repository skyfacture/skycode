param(
    [string]$WorkspaceRoot = "c:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode"
)

$ErrorActionPreference = "Stop"

$srcChm = Join-Path $WorkspaceRoot "swcomposerapi.chm"
$tmpRoot = "C:\temp\swcomposerapi_extract"
$tmpChm = Join-Path $tmpRoot "swcomposerapi.chm"
$tmpOut = Join-Path $tmpRoot "html"
$dstOut = Join-Path $WorkspaceRoot "docs\swcomposerapi_html"

if (-not (Test-Path -LiteralPath $srcChm)) {
    throw "CHM not found: $srcChm"
}

New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
New-Item -ItemType Directory -Path $tmpOut -Force | Out-Null
Copy-Item -LiteralPath $srcChm -Destination $tmpChm -Force
Unblock-File -LiteralPath $tmpChm -ErrorAction SilentlyContinue

& hh.exe -decompile "$tmpOut" "$tmpChm"

$maxAttempts = 20
$attempt = 0
$fileCount = 0
while ($attempt -lt $maxAttempts) {
    $fileCount = (Get-ChildItem -Path $tmpOut -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
    if ($fileCount -gt 0) { break }
    $attempt++
}

if ($fileCount -eq 0) {
    throw "Extraction appears empty. Check whether CHM is blocked/corrupt or HH is restricted."
}

New-Item -ItemType Directory -Path $dstOut -Force | Out-Null
Copy-Item -Path (Join-Path $tmpOut '*') -Destination $dstOut -Recurse -Force

$htmlFiles = Get-ChildItem -Path $dstOut -Recurse -File | Where-Object { $_.Extension -in @('.htm', '.html') } | Sort-Object FullName
$entries = @()

foreach ($f in $htmlFiles) {
    $raw = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction SilentlyContinue
    $title = ""

    if ($raw -match '(?is)<title>(.*?)</title>') {
        $title = ($Matches[1] -replace '\s+', ' ').Trim()
    }

    if ([string]::IsNullOrWhiteSpace($title) -and $raw -match '(?is)<h1[^>]*>(.*?)</h1>') {
        $title = ($Matches[1] -replace '<[^>]+>', '' -replace '\s+', ' ').Trim()
    }

    $text = ($raw -replace '(?is)<script.*?</script>', ' ' -replace '(?is)<style.*?</style>', ' ' -replace '(?is)<[^>]+>', ' ' -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '\s+', ' ').Trim()
    $tokens = [regex]::Matches($text, '\b[A-Za-z_][A-Za-z0-9_]{2,}\b') | ForEach-Object { $_.Value }
    $top = $tokens | Group-Object | Sort-Object Count -Descending | Select-Object -First 12 | ForEach-Object { $_.Name }

    $rel = Resolve-Path -LiteralPath $f.FullName -Relative
    if ($rel.StartsWith('.\')) { $rel = $rel.Substring(2) }

    $entries += [pscustomobject]@{
        file = ($rel -replace '\\', '/')
        title = $title
        topTerms = $top
    }
}

$json = [pscustomobject]@{
    generatedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ssK')
    source = 'swcomposerapi.chm'
    topicCount = $entries.Count
    topics = $entries
} | ConvertTo-Json -Depth 6
Set-Content -LiteralPath (Join-Path $WorkspaceRoot "COMPOSER_API_TOPIC_INDEX.json") -Value $json -Encoding UTF8

$md = @()
$md += '# Composer API Topic Index'
$md += ''
$md += ('Generated: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))
$md += ('Topics: ' + $entries.Count)
$md += ('Source: swcomposerapi.chm (extracted to docs/swcomposerapi_html)')
$md += ''
$md += '## Agent Usage'
$md += '- Search this file by API class, method, or enum keyword.'
$md += '- Open the matching extracted HTML topic under docs/swcomposerapi_html.'
$md += '- Use COMPOSER_API_TOPIC_INDEX.json for tool-based lookups.'
$md += ''

foreach ($e in $entries) {
    $line = '- ' + $e.file
    if (-not [string]::IsNullOrWhiteSpace($e.title)) { $line += ' | ' + $e.title }
    if ($e.topTerms.Count -gt 0) { $line += ' | terms: ' + (($e.topTerms -join ', ')) }
    $md += $line
}

Set-Content -LiteralPath (Join-Path $WorkspaceRoot "COMPOSER_API_TOPIC_INDEX.md") -Value ($md -join "`r`n") -Encoding UTF8

Write-Output ("Extraction files: " + ((Get-ChildItem -Path $dstOut -Recurse -File | Measure-Object).Count))
Write-Output ("Indexed topics: " + $entries.Count)
