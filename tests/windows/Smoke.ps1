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
    # A tall page must fill the width instead of shrinking to the available height.
    $tallPath=Join-Path $temp 'portrait.png'
    $tall=[Drawing.Bitmap]::new(1200,3600)
    $tallGraphics=[Drawing.Graphics]::FromImage($tall);$tallGraphics.Clear([Drawing.Color]::White)
    $tallFont=[Drawing.Font]::new('Arial',32)
    $tallGraphics.DrawString('FACTURE - DOCUMENT LONG', $tallFont, [Drawing.Brushes]::Black, 70, 130)
    $tallGraphics.DrawString('Reference FA-2026-0142', $tallFont, [Drawing.Brushes]::Black, 70, 240)
    $tallGraphics.DrawString('Client : exemple de controle', $tallFont, [Drawing.Brushes]::Black, 70, 350)
    for($line=0;$line -lt 12;$line++) { $tallGraphics.DrawString(('Prestation '+($line+1)+'                          125,00'),$tallFont,[Drawing.Brushes]::Black,70,(560+$line*120)) }
    $tallFont.Dispose();$tallGraphics.Dispose();$tall.Save($tallPath);$tall.Dispose()
    $canvas.LoadDocument($tallPath);$canvas.PerformLayout()
    $tallPicture=$type.GetField('picture',$flags).GetValue($canvas)
    if ($tallPicture.Width -lt 650 -or $tallPicture.Height -le 600) { throw 'Width fitting shrank a tall document to its height.' }
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
            $annotationDocument=New-Object Doctracker.Core.Models.DocumentRecord
            $annotationDocument.TestReference='DAC B 30 040';$annotationDocument.ReferenceNumber=1
            $annotation=New-Object Doctracker.Core.Models.DocumentComment
            $annotation.PageNumber=1;$annotation.X=.08;$annotation.Y=.68;$annotation.Width=.52;$annotation.Height=.25
            $annotation.Text='Pièce contrôlée et rapprochée.';$annotationDocument.Comments.Add($annotation)
            $viewCanvas.SetDocument($annotationDocument)
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
            if($documents.Bottom -gt $header.ClientSize.Height -or $documents.Top -lt 0) { throw 'Document selector clipped.' }
            if($header.Height -gt 60*$scenario.Scale) { throw 'Compact header uses too much height.' }
            foreach ($name in @('Documents','Categories','Query','Search','ModeState','Status','IndexState','Proofs')) {
                $control=$viewType.GetField($name,$flags).GetValue($view)
                if ($control -is [Windows.Forms.ComboBox] -and $control.ItemHeight -lt $control.Font.Height + 4) { throw "Native combo text clipped: $name" }
                $preferred=$control.GetPreferredSize([Drawing.Size]::new($control.Width,0))
                if ($control.Height + 2 -lt $preferred.Height) { throw "Clipped $name at $($scenario.Width) / $($scenario.Scale): $($control.Height) < $($preferred.Height)" }
                if ($control.Right -gt $control.Parent.ClientSize.Width + 2) { throw "Horizontal overflow: $name" }
            }
            $view.HideResults();$view.ShowProofs($false);$view.SetMode($null);$view.PerformLayout();[Windows.Forms.Application]::DoEvents()
            if($viewCanvas.Height -lt $view.Height*.70) { throw 'Less than 70 percent of the pane is available to the document.' }
            $view.ShowResults(0);$view.PerformLayout();[Windows.Forms.Application]::DoEvents()
            $emptyResults=$viewType.GetField('resultCount',$flags).GetValue($view)
            if(!$emptyResults.Visible -or $emptyResults.Text -notmatch 'Aucun résultat') { throw 'No visible empty-search feedback.' }
            if ($viewCanvas.Height -lt 160) { throw "Document area collapsed: $($viewCanvas.Height) at $($view.Width)x$($view.Height), scale $($scenario.Scale)" }

        } finally { $form.Dispose() }
    }
    # Actual mouse events on the rendered canvas: move and resize preview, Escape rollback.
    $form=[Windows.Forms.Form]::new();$form.ClientSize=[Drawing.Size]::new(800,600)
    $editCanvas=[Activator]::CreateInstance($type,$true);$form.Controls.Add($editCanvas);$form.Show()
    try {
        $editCanvas.LoadDocument($imagePath)
        $doc=New-Object Doctracker.Core.Models.DocumentRecord
        $comment=New-Object Doctracker.Core.Models.DocumentComment
        $comment.PageNumber=1;$comment.X=.1;$comment.Y=.2;$comment.Width=.3;$comment.Height=.2;$comment.Text='Texte test'
        $doc.Comments.Add($comment);$editCanvas.SetDocument($doc);[Windows.Forms.Application]::DoEvents()
        $editPicture=$type.GetField('picture',$flags).GetValue($editCanvas)
        $script:changedZone=$null
        $handler=[Action[string,Drawing.RectangleF]]{param($id,$zone) $script:changedZone=$zone}
        $editCanvas.add_CommentGeometryChanged($handler)
        $x=[int]($editPicture.Width*.2);$y=[int]($editPicture.Height*.3)
        $down=[Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,$x,$y,0)
        $move=[Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,$x+30,$y+20,0)
        $type.GetMethod('Picture_MouseDown',$flags).Invoke($editCanvas,@($editPicture,$down)) | Out-Null
        $type.GetMethod('Picture_MouseMove',$flags).Invoke($editCanvas,@($editPicture,$move)) | Out-Null
        $type.GetMethod('Picture_MouseUp',$flags).Invoke($editCanvas,@($editPicture,$move)) | Out-Null
        if($null -eq $script:changedZone -or $script:changedZone.X -le .1 -or [Math]::Abs($comment.X-.1) -gt .001) { throw 'Comment drag did not emit a move without mutating saved metadata.' }
        $script:changedZone=$null
        $x=[int]($editPicture.Width*.4);$y=[int]($editPicture.Height*.4)
        $down=[Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,$x,$y,0)
        $move=[Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,$x+40,$y+30,0)
        $type.GetMethod('Picture_MouseDown',$flags).Invoke($editCanvas,@($editPicture,$down)) | Out-Null
        $type.GetMethod('Picture_MouseMove',$flags).Invoke($editCanvas,@($editPicture,$move)) | Out-Null
        $type.GetMethod('Picture_MouseUp',$flags).Invoke($editCanvas,@($editPicture,$move)) | Out-Null
        if($null -eq $script:changedZone -or $script:changedZone.Width -le .3) { throw 'Comment resize handle failed.' }
        $script:changedZone=$null
        $type.GetMethod('Picture_MouseDown',$flags).Invoke($editCanvas,@($editPicture,$down)) | Out-Null
        $type.GetMethod('Picture_MouseMove',$flags).Invoke($editCanvas,@($editPicture,$move)) | Out-Null
        $type.GetMethod('CancelCommentDrag',$flags).Invoke($editCanvas,@()) | Out-Null
        $type.GetMethod('Picture_MouseUp',$flags).Invoke($editCanvas,@($editPicture,$move)) | Out-Null
        if($null -ne $script:changedZone) { throw 'Cancelled comment gesture was committed.' }
        $shot=[Drawing.Bitmap]::new($editCanvas.Width,$editCanvas.Height);$editCanvas.DrawToBitmap($shot,[Drawing.Rectangle]::new(0,0,$shot.Width,$shot.Height))
        $shot.Save((Join-Path $previewDirectory 'comment-handles.png'));$shot.Dispose()
    } finally {$form.Dispose()}
    Write-Host 'PASS: comment move, resize handles and cancellation via real mouse handlers'

    $folderStore=New-Object Doctracker.Core.Services.ProjectStore (Join-Path $temp 'folders')
    $folderState=New-Object Doctracker.Core.Models.ProjectState
    $folderService=New-Object Doctracker.Core.Services.ProjectFolders $folderStore
    $folderService.Create($folderState,'','Client 1');$folderService.Create($folderState,'Client 1','BL');$folderService.Create($folderState,'Client 1','Factures')
    foreach($name in @('BL-001.pdf','Facture-001.pdf')) {$item=New-Object Doctracker.Core.Models.DocumentRecord;$item.OriginalName=$name;$folderState.Documents.Add($item)}
    $folderType=$assembly.GetType('Doctracker.AddIn.UI.FolderOrganizer',$true)
    $organizer=$folderType.GetConstructors($flags)[0].Invoke(@($folderStore.PSObject.BaseObject,$folderState.PSObject.BaseObject))
    try {
        $organizer.Show();[Windows.Forms.Application]::DoEvents()
        $tree=$folderType.GetField('tree',$flags).GetValue($organizer)
        $files=$folderType.GetField('files',$flags).GetValue($organizer)
        if($files.Items.Count -ne 2 -or !$tree.AllowDrop) { throw 'Folder organizer did not show documents or allow drop.' }
        $folderType.GetMethod('Move',$flags).Invoke($organizer,[object[]]@([string[]]@($folderState.Documents[0].Id),'Client 1 / BL')) | Out-Null
        if($folderState.Documents[0].Categories[0] -ne 'Client 1 / BL' -or $files.Items.Count -ne 1) { throw 'Folder move failed to update membership and list.' }
        $shot=[Drawing.Bitmap]::new($organizer.Width,$organizer.Height);$organizer.DrawToBitmap($shot,[Drawing.Rectangle]::new(0,0,$shot.Width,$shot.Height))
        $shot.Save((Join-Path $previewDirectory 'folders.png'));$shot.Dispose()
    } finally {$organizer.Dispose()}
    Write-Host 'PASS: folder organizer, destination tree, document movement and refreshed list'

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
    $indexer.Index($state, $document, $null, [Threading.CancellationToken]::None, $false)
    if (!$document.IndexComplete -or $document.IndexedPages[0].Text -notmatch 'FA-001') { throw 'Native PDF index failed.' }
    $matcher = New-Object Doctracker.Core.Services.DocumentMatcher
    $hit = $matcher.Find($state, 'FA-001', 1)[0]
    if (!$hit.IsExact -or !$hit.HasLocation -or $hit.Y -gt .4 -or $hit.Y -lt .1) { throw 'Native PDF word position incorrect.' }
    [IO.File]::WriteAllBytes($store.IndexPath($document.IndexKey),[byte[]]@(1,2,3))
    $freshStore=New-Object Doctracker.Core.Services.ProjectStore $store.ProjectDirectory
    $freshState=$freshStore.LoadOrCreate('')
    $freshIndexer=$constructor.Invoke([object[]]@($freshStore.PSObject.BaseObject,$ocr.PSObject.BaseObject))
    $indexErrors=$freshIndexer.IndexMissing($freshState,$null,[Threading.CancellationToken]::None,$null,$true,$false)
    if($indexErrors.Count -ne 0 -or !$freshState.Documents[0].IndexComplete) { throw 'Damaged index did not rebuild from its source.' }
    $occurrences=[Doctracker.Core.Services.OccurrenceSearch]::Find($freshState,'001',10,[Threading.CancellationToken]::None)
    if($occurrences.Count -lt 1) { throw 'Search after automatic index recovery failed.' }
    # Failed files outside the active folder must not block search or trigger more disk writes.
    $badDoc=New-Object Doctracker.Core.Models.DocumentRecord
    $badDoc.RelativePath='documents/missing.pdf';$badDoc.OriginalName='missing.pdf';$badDoc.IndexError='Previous import failed'
    $freshState.Documents.Add($badDoc)
    $revisionBefore=$freshState.Revision
    $activeScope=[Collections.Generic.List[Doctracker.Core.Models.DocumentRecord]]::new()
    $activeScope.Add($freshState.Documents[0])
    $scopeErrors=$freshIndexer.IndexMissing($freshState,$null,[Threading.CancellationToken]::None,$activeScope,$false,$false)
    if($scopeErrors.Count -ne 0){throw 'Out-of-scope failed document blocked active folder.'}
    $knownErrors=$freshIndexer.IndexMissing($freshState,$null,[Threading.CancellationToken]::None,$null,$false,$false)
    if($knownErrors.Count -ne 1 -or $freshState.Revision -ne $revisionBefore){throw 'Known failed document was retried or caused an unnecessary save.'}
    # Cancellation before forced reindexing preserves the existing usable index.
    try { $freshIndexer.IndexMissing($freshState,$null,[Threading.CancellationToken]::new($true),$activeScope,$true,$true);throw 'Cancellation ignored' }
    catch { if($_.Exception.ToString() -notmatch 'OperationCanceledException'){throw} }
    if(!$freshState.Documents[0].IndexComplete){throw 'Cancelled reindex invalidated previous index.'}
    Write-Host 'PASS: folder-scoped indexing, no repeated failed imports, cancellation preserves previous index'
    Write-Host 'PASS: damaged XML/GZip index rebuilt; occurrence search works after reopening'
    Write-Host 'PASS: PDFium deployment, landscape rendering, native PDF text and positional matching'
    # Windows PowerShell does not apply the add-in's .dll.config redirects.
    # Exercise export in a tiny host with the deployed configuration, in each architecture.
    $exportPath=Join-Path $temp 'annotated.pdf'
    $exportHost=Join-Path $root 'Doctracker.ExportSmoke.exe'
    $platform=if ([IntPtr]::Size -eq 4) { 'x86' } else { 'x64' }
    $compiler=Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
    $exportSource=(Resolve-Path (Join-Path $PSScriptRoot 'ExportSmoke.cs')).Path
    $coreReference=Join-Path $root 'Doctracker.Core.dll'
    try {
        & $compiler /nologo /target:exe "/platform:$platform" "/out:$exportHost" "/reference:$coreReference" $exportSource
        if ($LASTEXITCODE -ne 0) { throw 'Export smoke host compilation failed.' }
        Copy-Item (Join-Path $root 'Doctracker.AddIn.dll.config') ($exportHost+'.config')
        $invariant=[Globalization.CultureInfo]::InvariantCulture
        $coordinates=@($hit.X,$hit.Y,$hit.Width,$hit.Height) | ForEach-Object { $_.ToString($invariant) }
        & $exportHost $pdfPath $exportPath @coordinates
        if ($LASTEXITCODE -ne 0) { throw 'Annotated export or its deployed dependency configuration failed.' }
    } finally {
        Remove-Item $exportHost,($exportHost+'.config') -ErrorAction SilentlyContinue
    }
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
