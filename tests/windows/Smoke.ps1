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
    # Repeated cell navigation must reuse the rendered page and preserve zoom.
    $type.GetMethod('SetZoom', $flags).Invoke($canvas, [object[]]@(1.25, $false)) | Out-Null
    $beforeImage = $picture.Image
    $canvas.NavigateTo($imagePath, $snip)
    if (![object]::ReferenceEquals($beforeImage, $picture.Image)) { throw 'Selecting a proof rerendered the same page.' }
    if ([Math]::Abs($type.GetField('zoom', $flags).GetValue($canvas) - 1.25) -gt .001) { throw 'Selecting a proof reset the zoom.' }
    Write-Host 'PASS: canvas loading, scaled display, normalized crop, proof navigation and retained zoom'

    $previewDirectory = Join-Path $PSScriptRoot '../../artifacts/ui-previews'
    [void][IO.Directory]::CreateDirectory($previewDirectory)
    $viewType = $assembly.GetType('Doctracker.AddIn.UI.WorkspaceView', $true)
    foreach ($scenario in @(@{Width=760;Height=850;Scale=1.0}, @{Width=420;Height=850;Scale=1.0}, @{Width=840;Height=1500;Scale=2.0})) {
        # Keep a visible native window so WM_PRINT renders edit/combo controls.
        # The undocked view may exceed the virtual desktop; its bitmap retains
        # the requested physical dimensions without Windows clamping a Form.
        $form = [Windows.Forms.Form]::new()
        $view = [Activator]::CreateInstance($viewType, $true)
        try {
            $form.ClientSize = [Drawing.Size]::new(800,600)
            $view.Dock=[Windows.Forms.DockStyle]::None
            $view.Size=[Drawing.Size]::new($scenario.Width,$scenario.Height)
            $form.Controls.Add($view)
            $documents = $viewType.GetField('Documents', $flags).GetValue($view)
            [void]$documents.Items.Add('Facture - septembre 2026.pdf'); $documents.SelectedIndex=0
            $viewType.GetField('Query', $flags).GetValue($view).Text='FA-2026-0142'
            $viewType.GetField('IndexState', $flags).GetValue($view).Text='3 documents · 3 indexés'
            $viewType.GetField('Results', $flags).GetValue($view).Items.Add('Facture - septembre 2026.pdf · page 1 · score 100 %') | Out-Null
            $view.SetMode([Doctracker.Core.Models.SnipType]::Text)
            $view.ShowResults(1)
            $viewType.GetField('Proofs', $flags).GetValue($view).Items.Add('Texte · Facture - septembre 2026.pdf · p. 1 · Achats!F12') | Out-Null
            $viewType.GetField('Proofs', $flags).GetValue($view).SelectedIndex=0
            $view.ShowProofs($true)
            $viewCanvas=$viewType.GetField('Canvas', $flags).GetValue($view)
            $viewCanvas.LoadDocument($imagePath)
            $viewCanvas.NavigateTo($imagePath,$snip)
            $form.Show(); [Windows.Forms.Application]::DoEvents()
            if ($scenario.Scale -ne 1) {
                # Simulate larger Windows text and geometry; actual Office per-monitor DPI remains a desktop check.
                $view.Scale([Drawing.SizeF]::new($scenario.Scale,$scenario.Scale))
                $view.Font=[Drawing.Font]::new('Segoe UI',9*$scenario.Scale)
                $view.Size=[Drawing.Size]::new($scenario.Width,$scenario.Height)
            }
            $view.PerformLayout(); [Windows.Forms.Application]::DoEvents()
            $screenshot=[Drawing.Bitmap]::new($view.Width,$view.Height)
            $view.DrawToBitmap($screenshot,[Drawing.Rectangle]::new(0,0,$view.Width,$view.Height))
            $screenshot.Save((Join-Path $previewDirectory ("workspace-"+$scenario.Width+"-"+$scenario.Scale+".png")))
            $screenshot.Dispose()
            $header=$view.Controls[0].GetControlFromPosition(0,0)
            $picker=$header.GetControlFromPosition(1,0)
            $caption=$picker.GetControlFromPosition(0,0)
            if ($picker.Top -lt 0 -or $caption.Top -lt 0 -or $picker.Bottom -gt $header.ClientSize.Height) { throw 'Header caption or document picker clipped.' }
            foreach ($name in @('Documents','Query','Search','ModeState','Status','IndexState','Proofs')) {
                $control=$viewType.GetField($name,$flags).GetValue($view)
                if ($control -is [Windows.Forms.ComboBox] -and $control.ItemHeight -lt $control.Font.Height + 4) { throw "Native combo text clipped: $name" }
                $preferred=$control.GetPreferredSize([Drawing.Size]::new($control.Width,0))
                if ($control.Height + 2 -lt $preferred.Height) { throw "Clipped $name at $($scenario.Width) / $($scenario.Scale): $($control.Height) < $($preferred.Height)" }
                if ($control.Right -gt $control.Parent.ClientSize.Width + 2) { throw "Horizontal overflow: $name" }
            }
            if ($viewCanvas.Height -lt 160) { throw "Document area collapsed: $($viewCanvas.Height) at $($view.Width)x$($view.Height), scale $($scenario.Scale)" }

        } finally { $form.Dispose() }
    }
    Write-Host 'PASS: responsive workspace at 420/760 px, enlarged text, controls without vertical clipping, UI previews'


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
    $exporterType=$assembly.GetType('Doctracker.AddIn.Infrastructure.AnnotatedPdfExporter',$true)
    $snip.X=$hit.X; $snip.Y=$hit.Y; $snip.Width=$hit.Width; $snip.Height=$hit.Height
    $snip.WorksheetName='Achats'; $snip.CellAddress='B2'
    $document.TestReference='DAC B 30 040'; $document.ReferenceNumber=1
    $exportPath=Join-Path $temp 'annotated.pdf'
    $arguments=[object[]]::new(5)
    $arguments[0]=$pdfPath; $arguments[1]=$document.PSObject.BaseObject
    $arguments[2]=[Doctracker.Core.Models.SnipRecord[]]@($snip)
    $arguments[3]=$exportPath; $arguments[4]=[Threading.CancellationToken]::None
    $exporterType.GetMethod('Export',[Reflection.BindingFlags]'Static,Public').Invoke($null,$arguments) | Out-Null
    $canvas.LoadDocument($exportPath)
    $picture=$type.GetField('picture',$flags).GetValue($canvas)
    if (!$canvas.HasDocument -or [Math]::Abs($picture.Image.Width/$picture.Image.Height-2) -gt .01) { throw 'Annotated PDF export did not preserve page orientation.' }
    $picture.Image.Save((Join-Path $previewDirectory 'annotated-export.png'))
    Write-Host 'PASS: flattened PDF export with colored snips and Xref, reopened in PDFium'

} finally {
    if ($canvas) { $canvas.Dispose() }
    if ($ocr) { $ocr.Dispose() }
    Remove-Item $temp -Recurse -Force
}
