[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'

$inputFile = Get-Item -LiteralPath $InputPath
if ($inputFile.PSIsContainer) {
    throw "InputPath must be an MHT file, not a directory."
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = [IO.Path]::ChangeExtension($inputFile.FullName, '.pdf')
}
$OutputPath = [IO.Path]::GetFullPath($OutputPath)

if ([string]::Equals($inputFile.FullName, $OutputPath, [StringComparison]::OrdinalIgnoreCase)) {
    throw "The PDF output path must not overwrite the source MHT."
}

$outputDirectory = [IO.Path]::GetDirectoryName($OutputPath)
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    throw "Output directory does not exist: $outputDirectory"
}

$tempPath = Join-Path $outputDirectory ([IO.Path]::GetRandomFileName() + '.pdf')
$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputFile.FullName, $false, $true)

    foreach ($section in $document.Sections) {
        $page = $section.PageSetup
        $page.PaperSize = 7
        $page.Orientation = 0
        $page.LeftMargin = 18
        $page.RightMargin = 18
        $page.TopMargin = 18
        $page.BottomMargin = 18
    }

    $page = $document.Sections.Item(1).PageSetup
    $maxWidth = $page.PageWidth - $page.LeftMargin - $page.RightMargin
    $maxHeight = $page.PageHeight - $page.TopMargin - $page.BottomMargin

    for ($index = 1; $index -le $document.InlineShapes.Count; $index++) {
        $image = $document.InlineShapes.Item($index)
        $scale = [Math]::Min(1.0, [Math]::Min($maxWidth / $image.Width, $maxHeight / $image.Height))
        if ($scale -lt 1.0) {
            $image.LockAspectRatio = -1
            $image.Width = $image.Width * $scale
        }
    }

    $document.Repaginate()
    $expectedPages = $document.ComputeStatistics(2)
    if ($expectedPages -lt 1) {
        throw "Word did not produce any pages."
    }

    $document.ExportAsFixedFormat($tempPath, 17)
}
finally {
    if ($null -ne $document) {
        $document.Close(0)
    }
    if ($null -ne $word) {
        $word.Quit()
    }
    if ($null -ne $document) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    if ($null -ne $word) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}

try {
    if (-not (Test-Path -LiteralPath $tempPath -PathType Leaf)) {
        throw "Word did not create the PDF."
    }

    $bytes = [IO.File]::ReadAllBytes($tempPath)
    if ($bytes.Length -lt 5 -or [Text.Encoding]::ASCII.GetString($bytes, 0, 5) -ne '%PDF-') {
        throw "The exported file is not a valid PDF."
    }

    $pdf = [Text.Encoding]::ASCII.GetString($bytes)
    $actualPages = [regex]::Matches($pdf, '/Type\s*/Page\b').Count
    $pageBoxes = [regex]::Matches($pdf, '/MediaBox\s*\[([^\]]+)\]')
    if ($actualPages -ne $expectedPages -or $pageBoxes.Count -ne $actualPages) {
        throw "PDF page verification failed: Word reported $expectedPages pages; the PDF contains $actualPages pages and $($pageBoxes.Count) page boxes."
    }

    foreach ($box in $pageBoxes) {
        $coordinates = $box.Groups[1].Value.Trim() -split '\s+'
        if ($coordinates.Count -ne 4) {
            throw "PDF contains an unrecognized page box."
        }

        $width = [double]::Parse($coordinates[2], [Globalization.CultureInfo]::InvariantCulture) -
            [double]::Parse($coordinates[0], [Globalization.CultureInfo]::InvariantCulture)
        $height = [double]::Parse($coordinates[3], [Globalization.CultureInfo]::InvariantCulture) -
            [double]::Parse($coordinates[1], [Globalization.CultureInfo]::InvariantCulture)
        if ([Math]::Abs($width - 595.28) -gt 1 -or [Math]::Abs($height - 841.89) -gt 1) {
            throw "PDF contains a page that is not portrait A4."
        }
    }

    Move-Item -LiteralPath $tempPath -Destination $OutputPath -Force
    Write-Output "Created $OutputPath ($actualPages portrait A4 pages)."
}
finally {
    if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
        Remove-Item -LiteralPath $tempPath
    }
}
