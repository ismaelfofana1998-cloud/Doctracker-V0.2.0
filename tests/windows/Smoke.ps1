param([string]$BuildDirectory = "$PSScriptRoot\..\..\src\Doctracker.AddIn\bin\Release")
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path $BuildDirectory).Path
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void][Reflection.Assembly]::LoadFrom((Join-Path $root 'Doctracker.Core.dll'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $root 'PdfiumViewer.dll'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $root 'Tesseract.dll'))
$assembly = [Reflection.Assembly]::LoadFrom((Join-Path $root 'Doctracker.AddIn.dll'))
$flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
$temp = Join-Path ([IO.Path]::GetTempPath()) ('doctracker-smoke-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
$canvas = $null; $ocr = $null
try {
    $type = $assembly.GetType('Doctracker.AddIn.UI.DocumentCanvas', $true)
    $canvas = [Activator]::CreateInstance($type, $true)
    $canvas.Size = [Drawing.Size]::new(760, 600)
    $canvas.CreateControl()
    $imagePath = Join-Path $temp 'scan.png'
    $bitmap = [Drawing.Bitmap]::new(1200, 600)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([Drawing.Color]::White)
    $font = [Drawing.Font]::new('Arial', 40)
    $graphics.DrawString('FACTURE 12345', $font, [Drawing.Brushes]::Black, 100, 150)
    $graphics.FillRectangle([Drawing.Brushes]::Red, 900, 400, 200, 100)
    $bitmap.Save($imagePath)
    $font.Dispose(); $graphics.Dispose(); $bitmap.Dispose()
    $canvas.LoadDocument($imagePath)
    $picture = $type.GetField('picture', $flags).GetValue($canvas)
    if ($picture.SizeMode -ne [Windows.Forms.PictureBoxSizeMode]::StretchImage) { throw 'Zoom does not scale the source image.' }
    $snip = New-Object Doctracker.Core.Models.SnipRecord
    $snip.PageNumber = 1; $snip.X = .75; $snip.Y = 0.6666666667; $snip.Width = 0.16; $snip.Height = 0.16
    $canvas.NavigateTo($imagePath, $snip)
    $crop = $canvas.CropSelection()
    if ($crop.GetPixel(30, 30).R -lt 240 -or $crop.GetPixel(30, 30).G -gt 30) { throw 'Crop coordinates drifted from the selected region.' }
    $crop.Dispose()
    Write-Host 'PASS: canvas loading, scaled display, normalized crop and proof navigation'

    $ocrType = $assembly.GetType('Doctracker.AddIn.Infrastructure.TesseractOcrEngine', $true)
    $ocr = [Activator]::CreateInstance($ocrType, $true)
    $bitmap = [Drawing.Bitmap]::new($imagePath)
    $page = $ocr.Recognize($bitmap, $false)
    $bitmap.Dispose()
    if ($page.Text -notmatch '12345' -or $page.Words.Count -lt 2) { throw ('OCR text/word boxes missing: ' + $page.Text) }
    Write-Host 'PASS: deployed native OCR, French/English models, recognized text and word boxes'

    # Minimal PDF fixture built with exact byte offsets; no external files or customer data.
    $objects = @(
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
    )
    $stream = 'BT /F1 24 Tf 60 220 Td (FACTURE FA-001 TOTAL 1250.00) Tj ET'
    $objects += "<< /Length $($stream.Length) >>`nstream`n$stream`nendstream"
    $pdfText = "%PDF-1.4`n"
    $offsets = [Collections.Generic.List[int]]::new()
    for ($i = 0; $i -lt $objects.Count; $i++) {
        $offsets.Add($pdfText.Length)
        $pdfText += "$($i + 1) 0 obj`n$($objects[$i])`nendobj`n"
    }
    $xref = $pdfText.Length
    $pdfText += "xref`n0 6`n0000000000 65535 f `n"
    foreach ($offset in $offsets) { $pdfText += ('{0:D10} 00000 n ' -f $offset) + "`n" }
    $pdfText += "trailer`n<< /Size 6 /Root 1 0 R >>`nstartxref`n$xref`n%%EOF"
    $pdfPath = Join-Path $temp 'native.pdf'
    [IO.File]::WriteAllText($pdfPath, $pdfText, [Text.Encoding]::ASCII)
    $canvas.LoadDocument($pdfPath)
    $picture = $type.GetField('picture', $flags).GetValue($canvas)
    if ([Math]::Abs($picture.Image.Width / $picture.Image.Height - 2) -gt .01) { throw 'Landscape PDF aspect ratio was distorted.' }
    $store = New-Object Doctracker.Core.Services.ProjectStore (Join-Path $temp 'project')
    $state = New-Object Doctracker.Core.Models.ProjectState
    $importer = New-Object Doctracker.Core.Services.DocumentImporter $store
    $document = $importer.Import($state, $pdfPath, 'smoke')
    $indexerType = $assembly.GetType('Doctracker.AddIn.Infrastructure.DocumentIndexer', $true)
    $constructor = $indexerType.GetConstructors($flags)[0]
    $indexer = $constructor.Invoke([object[]]@($store.PSObject.BaseObject, $ocr.PSObject.BaseObject))
    $indexer.Index($state, $document, $null, [Threading.CancellationToken]::None)
    if (!$document.IndexComplete -or $document.IndexedPages[0].Text -notmatch 'FA-001') { throw 'Native PDF index failed.' }
    $matcher = New-Object Doctracker.Core.Services.DocumentMatcher
    $hit = $matcher.Find($state, 'FA-001', 1)[0]
    if (!$hit.IsExact -or !$hit.HasLocation -or $hit.Y -gt .4 -or $hit.Y -lt .1) { throw 'Native PDF word position incorrect.' }
    Write-Host 'PASS: PDFium deployment, landscape rendering, native PDF text and positional matching'
} finally {
    if ($canvas) { $canvas.Dispose() }
    if ($ocr) { $ocr.Dispose() }
    Remove-Item $temp -Recurse -Force
}
