param([string]$BuildDirectory = "$PSScriptRoot\..\..\src\Doctracker.AddIn\bin\Release")
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path $BuildDirectory).Path
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::SetUnhandledExceptionMode([Windows.Forms.UnhandledExceptionMode]::ThrowException)
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
    if($canvas.ActiveType -ne [Doctracker.Core.Models.SnipType]::Text){throw 'Text snip is not the default.'}
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
            $view.SetDocumentSummary(3,0)
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
            if($documents.Bottom -gt $documents.Parent.ClientSize.Height -or $documents.Top -lt 0) { throw "Document selector clipped: top=$($documents.Top), bottom=$($documents.Bottom), parent=$($documents.Parent.ClientSize.Height), item=$($documents.ItemHeight), preferred=$($documents.PreferredHeight)." }
            if($header.Height -gt 60*$scenario.Scale) { throw 'Compact header uses too much height.' }
            foreach ($name in @('Brand','Documents','Categories','Query','Search','Proofs')) {
                $control=$viewType.GetField($name,$flags).GetValue($view)
                if ($control -is [Windows.Forms.ComboBox] -and $control.ItemHeight -lt $control.Font.Height + 4) { throw "Native combo text clipped: $name" }
                $preferred=$control.GetPreferredSize([Drawing.Size]::new($control.Width,0))
                if (!($control -is [Windows.Forms.ComboBox]) -and $control.Height + 2 -lt $preferred.Height) { throw "Clipped $name at $($scenario.Width) / $($scenario.Scale): $($control.Height) < $($preferred.Height)" }
                if ($control.Right -gt $control.Parent.ClientSize.Width + 2) { throw "Horizontal overflow: $name" }
            }
            $view.HideResults();$view.ShowProofs($false);$view.SetMode($null);$view.PerformLayout();[Windows.Forms.Application]::DoEvents()
            $categories=$viewType.GetField('Categories',$flags).GetValue($view)
            if($categories.Parent -ne $documents.Parent -or [Math]::Abs($categories.Top-$documents.Top) -gt 3){throw 'Folders and documents are not on the same row.'}
            $brand=$viewType.GetField('Brand',$flags).GetValue($view)
            $query=$viewType.GetField('Query',$flags).GetValue($view)
            if($brand.Parent -ne $query.Parent.Parent){throw 'Search is not on the Doctracker header row.'}
            $restHeight=$viewCanvas.Height
            foreach($mode in @([Doctracker.Core.Models.SnipType]::Text,[Doctracker.Core.Models.SnipType]::Exception)) {
                $view.SetMode($mode);$view.PerformLayout();[Windows.Forms.Application]::DoEvents()
                if($viewCanvas.Height -ne $restHeight){throw 'Snip mode added a banner or reduced the document area.'}
            }
            $view.SetCommentMode($true);$view.PerformLayout();[Windows.Forms.Application]::DoEvents()
            if($viewCanvas.Height -ne $restHeight){throw 'Comment mode reduced the document area.'}
            $view.SetMode($null)
            $viewType.GetField('Status',$flags).GetValue($view).Text=('Long status message ' * 30)
            $view.PerformLayout();[Windows.Forms.Application]::DoEvents()
            if($viewCanvas.Height -ne $restHeight){throw 'Long status expanded the footer.'}
            $viewType.GetField('Status',$flags).GetValue($view).Text='Prêt'
            $clean=[Drawing.Bitmap]::new($view.Width,$view.Height);$view.DrawToBitmap($clean,[Drawing.Rectangle]::new(0,0,$view.Width,$view.Height))
            $clean.Save((Join-Path $previewDirectory ('compact-'+$scenario.Width+'-'+$scenario.Scale+'.png')));$clean.Dispose()

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
        $focusProbe=[Windows.Forms.TextBox]::new();$form.Controls.Add($focusProbe);$focusProbe.BringToFront();$focusProbe.Focus() | Out-Null
        [Windows.Forms.Application]::DoEvents()
        if(!$focusProbe.Focused){throw 'Keyboard focus probe could not be established.'}
        $editPicture.GetType().GetMethod('OnMouseEnter',$flags).Invoke($editPicture,@([EventArgs]::Empty)) | Out-Null
        if(!$focusProbe.Focused){throw 'Hovering PDF stole keyboard focus.'}

        $script:changedZone=$null
        $handler=[Action[string,Drawing.RectangleF]]{param($id,$zone) $script:changedZone=$zone}
        $editCanvas.add_CommentGeometryChanged($handler)
        $x=[int]($editPicture.Width*.2);$y=[int]($editPicture.Height*.3)
        $down=[Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,$x,$y,0)
        $move=[Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,$x+30,$y+20,0)
        $type.GetMethod('Picture_MouseDown',$flags).Invoke($editCanvas,@($editPicture,$down)) | Out-Null
        $editViewport=$type.GetField('viewport',$flags).GetValue($editCanvas)
        if(!$editViewport.Focused){throw 'Clicking PDF did not enable Escape keyboard handling.'}
        $focusProbe.Dispose()

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


    $ocrType = $assembly.GetType('Doctracker.AddIn.Infrastructure.IsolatedOcrEngine', $true)
    $ocr = [Activator]::CreateInstance($ocrType, $true)
    $bitmap = [Drawing.Bitmap]::new($imagePath)
    $page = $ocr.Recognize($bitmap, $false)
    $bitmap.Dispose()
    if ($page.Text -notmatch '12345' -or $page.Words.Count -lt 2) { throw ('OCR text/word boxes missing: ' + $page.Text) }
    Write-Host 'PASS: deployed native OCR, French/English models, recognized text and word boxes'
    $ocrZone=[Drawing.RectangleF]::new(0,.1,.65,.55)
    $regionPage=$ocr.RecognizeRegion($imagePath,1,$ocrZone,$false,[Threading.CancellationToken]::None)
    if($regionPage.Text -notmatch '12345'){throw 'Isolated source-region OCR did not preserve crop coordinates.'}
    $badSource=Join-Path $temp 'broken.pdf';[IO.File]::WriteAllText($badSource,'invalid synthetic PDF')
    $failed=$false
    try {$ocr.RecognizeRegion($badSource,1,$ocrZone,$false,[Threading.CancellationToken]::None) | Out-Null}
    catch {$failed=$true}
    if(!$failed){throw 'Invalid PDF worker result was accepted.'}
    $page=$ocr.RecognizeRegion($imagePath,1,$ocrZone,$false,[Threading.CancellationToken]::None)
    if($page.Text -notmatch '12345'){throw 'OCR did not recover after a failed child.'}

    $blank=[Drawing.Bitmap]::new(200,100)
    $blankGraphics=[Drawing.Graphics]::FromImage($blank);$blankGraphics.Clear([Drawing.Color]::White);$blankGraphics.Dispose()
    try {$emptyPage=$ocr.Recognize($blank,$false);if($emptyPage.Text -ne '' -or $emptyPage.Words.Count -ne 0){throw 'Blank OCR should produce an empty result.'}}
    finally {$blank.Dispose()}
    Write-Host 'PASS: blank OCR without native word iteration'

    # Unexpected child exit and a hung child must not terminate the host.
    $crashWorker=Join-Path $temp 'fake-crash.exe'
    Add-Type -TypeDefinition 'public static class CrashChild {public static int Main(string[] args){System.Environment.Exit(139);return 139;}}' -OutputAssembly $crashWorker -OutputType ConsoleApplication
    $hangWorker=Join-Path $temp 'fake-hang.exe'
    Add-Type -TypeDefinition 'public static class HungChild {public static void Main(string[] args){System.IO.File.WriteAllText(System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory,"fake-worker.pid"),System.Diagnostics.Process.GetCurrentProcess().Id.ToString());System.Threading.Thread.Sleep(30000);}}' -OutputAssembly $hangWorker -OutputType ConsoleApplication
    $constructor=$ocrType.GetConstructor($flags,$null,[Type[]]@([string],[int]),$null)
    $workRoot=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Doctracker/OcrWork'
    $beforeWork=@(Get-ChildItem $workRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object Name)
    foreach($scenario in @(@{Path=$crashWorker;Timeout=3000;Cancel=$false},@{Path=$hangWorker;Timeout=1200;Cancel=$false},@{Path=$hangWorker;Timeout=5000;Cancel=$true})) {
        $isolated=$constructor.Invoke([object[]]@([string]$scenario.Path,[int]$scenario.Timeout))
        $cancel=[Threading.CancellationTokenSource]::new();$rejected=$false
        if($scenario.Cancel){$cancel.CancelAfter(500)}
        $elapsed=[Diagnostics.Stopwatch]::StartNew()
        try {$isolated.RecognizeRegion($imagePath,1,$ocrZone,$false,$cancel.Token) | Out-Null}
        catch {$rejected=$true}
        finally {$isolated.Dispose();$cancel.Dispose()}
        if(!$rejected -or $elapsed.Elapsed.TotalSeconds -gt 10){throw 'Child failure/timeout/cancellation was not contained promptly.'}
        if($scenario.Path -eq $hangWorker) {
            $pidFile=Join-Path $temp 'fake-worker.pid'
            if(Test-Path $pidFile){$childId=[int][IO.File]::ReadAllText($pidFile);if(Get-Process -Id $childId -ErrorAction SilentlyContinue){throw 'Cancelled/hung OCR child survived.'};Remove-Item $pidFile}
        }
    }
    $afterWork=@(Get-ChildItem $workRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object Name)
    if(@($afterWork | Where-Object {$_ -notin $beforeWork}).Count -ne 0){throw 'OCR temporary files leaked after a worker failure.'}
    Write-Host 'PASS: isolated OCR and crop, invalid source recovery, child exit, timeout, cancellation and temporary cleanup'
    # The same containment must hold for batch requests, with no partial result accepted.
    foreach($scenario in @(@{Path=$crashWorker;Timeout=3000;Cancel=$false},@{Path=$hangWorker;Timeout=1200;Cancel=$false},@{Path=$hangWorker;Timeout=5000;Cancel=$true})) {
        $isolated=$constructor.Invoke([object[]]@([string]$scenario.Path,[int]$scenario.Timeout))
        $cancel=[Threading.CancellationTokenSource]::new();$rejected=$false
        if($scenario.Cancel){$cancel.CancelAfter(500)}
        try {$isolated.RecognizePages($imagePath,[int[]]@(1),$null,$cancel.Token) | Out-Null}
        catch {$rejected=$true}
        finally {$isolated.Dispose();$cancel.Dispose()}
        if(!$rejected){throw 'Batch accepted failed or cancelled worker result.'}
        $pidFile=Join-Path $temp 'fake-worker.pid'
        if(Test-Path $pidFile){$childId=[int][IO.File]::ReadAllText($pidFile);if(Get-Process -Id $childId -ErrorAction SilentlyContinue){throw 'Batch left a child alive.'};Remove-Item $pidFile}
    }
    $heartbeatWorker=Join-Path $temp 'fake-heartbeat.exe'
    Add-Type -TypeDefinition @'
public static class HeartbeatChild {
    public static void Main(string[] args) {
        for(int i=1;i<=3;i++){System.Threading.Thread.Sleep(600);System.Console.WriteLine("WorkerPageReady "+i);}
        System.IO.File.WriteAllText(args[1],"<OcrBatchResult><Pages><PageTextRecord PageNumber='1'>one</PageTextRecord><PageTextRecord PageNumber='2'>two</PageTextRecord><PageTextRecord PageNumber='3'>three</PageTextRecord></Pages></OcrBatchResult>");
    }
}
'@ -OutputAssembly $heartbeatWorker -OutputType ConsoleApplication
    $isolated=$constructor.Invoke([object[]]@([string]$heartbeatWorker,[int]1400))
    try {$heartbeatPages=$isolated.RecognizePages($imagePath,[int[]]@(1,2,3),$null,[Threading.CancellationToken]::None);if($heartbeatPages.Count -ne 3){throw 'Heartbeat batch result missing.'}}
    finally {$isolated.Dispose()}
    if(@(Get-ChildItem $workRoot -Directory | Where-Object {$_.Name -notin $beforeWork}).Count -ne 0){throw 'Batch leaked temporary work folders.'}
    Write-Host 'PASS: batch crash, timeout, cancellation, per-page deadline and cleanup'



    # Minimal PDF fixture built with exact byte offsets; no external files or customer data.
    $objects = @(
        '<< /Type /Catalog /Pages 2 0 R >>',
        '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
    )
    $stream = 'BT /F1 24 Tf 60 220 Td (FACTURE FA-001 TOTAL 1250.00) Tj 0 -40 Td (REF FAC12345A) Tj 0 -40 Td [(5) 20 (0) 20 (0) 20 (X) 20 (3) 20 (0) 20 (0) 20 (Z) 20 (3) 20 (5)] TJ ET'
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
    $readPage=$indexerType.GetMethod('ReadNativePage',($flags -bor [Reflection.BindingFlags]::Static),$null,[Type[]]@([string],[int]),$null)
    $nativePage=$readPage.Invoke($null,[object[]]@($pdfPath.PSObject.BaseObject,1))
    if($document.IndexComplete -or $document.IndexKey -ne '' -or $nativePage.Text -notmatch 'FA-001' -or $nativePage.Words.Count -eq 0){throw 'First snip cannot use native PDF text before indexing.'}

    $recoveryCount=@(Get-ChildItem (Join-Path $store.ProjectDirectory 'recovery') -Filter *.xml).Count
    $indexer.Index($state, $document, $null, [Threading.CancellationToken]::None, $false, $true)
    if(@(Get-ChildItem (Join-Path $store.ProjectDirectory 'recovery') -Filter *.xml).Count -ne $recoveryCount){throw 'Indexing created redundant recovery snapshots.'}
    if (!$document.IndexComplete -or $document.IndexedPages[0].Text -notmatch 'FA-001') { throw 'Native PDF index failed.' }
    $matcher = New-Object Doctracker.Core.Services.DocumentMatcher
    $hit = $matcher.Find($state, 'FA-001', 1)[0]
    if (!$hit.IsExact -or !$hit.HasLocation -or $hit.Y -gt .4 -or $hit.Y -lt .1) { throw 'Native PDF word position incorrect.' }
    foreach($query in @('12345','X300','500X300Z35')) {
        $literal=[Doctracker.Core.Services.OccurrenceSearch]::Find($state,$query,10,[Threading.CancellationToken]::None)
        $matched=$matcher.Find($state,$query,10,$true,[Threading.CancellationToken]::None)
        if($literal.Count -eq 0 -or !$literal[0].HasLocation -or $matched.Count -eq 0 -or !$matched[0].HasLocation){throw "Native embedded reference not found with position: $query"}
    }
    Write-Host 'PASS: native PDF embedded numeric references and split glyphs, shared search/matching, no redundant recovery copy'
    Add-Type -Path (Join-Path $PSScriptRoot 'TableEditorProbe.cs') -ReferencedAssemblies (Join-Path $root 'Doctracker.Core.dll'),'System.Drawing.dll','System.Windows.Forms.dll','System.Core.dll'
    Write-Host ([TableEditorProbe]::Run($assembly,$pdfPath,$nativePage,$previewDirectory))

    [IO.File]::WriteAllBytes($store.IndexPath($document.IndexKey),[byte[]]@(1,2,3))
    $freshStore=New-Object Doctracker.Core.Services.ProjectStore $store.ProjectDirectory
    $freshState=$freshStore.LoadOrCreate('')
    $freshIndexer=$constructor.Invoke([object[]]@($freshStore.PSObject.BaseObject,$ocr.PSObject.BaseObject))
    $indexErrors=$freshIndexer.IndexMissing($freshState,$null,[Threading.CancellationToken]::None,$null,$true,$false,$true)
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
    $scopeErrors=$freshIndexer.IndexMissing($freshState,$null,[Threading.CancellationToken]::None,$activeScope,$false,$false,$true)
    if($scopeErrors.Count -ne 0){throw 'Out-of-scope failed document blocked active folder.'}
    $knownErrors=$freshIndexer.IndexMissing($freshState,$null,[Threading.CancellationToken]::None,$null,$false,$false,$true)
    if($knownErrors.Count -ne 1 -or $freshState.Revision -ne $revisionBefore){throw 'Known failed document was retried or caused an unnecessary save.'}
    # Cancellation before forced reindexing preserves the existing usable index.
    try { $freshIndexer.IndexMissing($freshState,$null,[Threading.CancellationToken]::new($true),$activeScope,$true,$true,$true);throw 'Cancellation ignored' }
    catch { if($_.Exception.ToString() -notmatch 'OperationCanceledException'){throw} }
    if(!$freshState.Documents[0].IndexComplete){throw 'Cancelled reindex invalidated previous index.'}
    Write-Host 'PASS: folder-scoped indexing, no repeated failed imports, cancellation preserves previous index'
    Write-Host 'PASS: damaged XML/GZip index rebuilt; occurrence search works after reopening'
    # Native text preparation never runs OCR; scanned images remain pending until allowed.
    $scanState=New-Object Doctracker.Core.Models.ProjectState
    $scanDoc=$importer.Import($scanState,$imagePath,'ocr-demand')
    $indexer.Index($scanState,$scanDoc,$null,[Threading.CancellationToken]::None,$false,$false)
    if($scanDoc.IndexComplete -or $scanDoc.IndexedPages[0].Text -ne ''){throw 'A scan was recognized without OCR authorization.'}
    $indexer.Index($scanState,$scanDoc,$null,[Threading.CancellationToken]::None,$false,$true)
    if(!$scanDoc.IndexComplete -or $scanDoc.IndexedPages[0].Text -notmatch '12345'){throw 'On-demand OCR did not complete.'}
    Write-Host 'PASS: scanned image remains pending until explicit OCR, then becomes searchable'
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

    $paneType=$assembly.GetType('Doctracker.AddIn.UI.DoctrackerPaneControl',$true)
    $sumSource=$paneType.GetMethod('SumSourceText',($flags -bor [Reflection.BindingFlags]::Static))
    $sumPage=New-Object Doctracker.Core.Models.PageTextRecord
    foreach($item in @(@{Text='10';X=.1},@{Text='200';X=.6})) {
        $word=New-Object Doctracker.Core.Models.WordRecord
        $word.Text=$item.Text;$word.X=$item.X;$word.Y=.1;$word.Width=.1;$word.Height=.03;$word.Line=1
        $sumPage.Words.Add($word)
    }
    $sumText=$sumSource.Invoke($null,[object[]]@($sumPage.PSObject.BaseObject))
    $parser=New-Object Doctracker.Core.Services.TextValueParser
    if($parser.Parse([Doctracker.Core.Models.SnipType]::Sum,$sumText) -ne '210'){throw 'Sum joined separate columns into a thousands group.'}
    Write-Host 'PASS: snip sum preserves separate numeric columns'
    # Office can call us without a UI SynchronizationContext: success, failure and
    # cancellation must all resume on the owner thread before touching GDI/Excel.
    Add-Type -Path (Join-Path $PSScriptRoot 'UiContinuationProbe.cs') -ReferencedAssemblies 'System.Windows.Forms.dll','System.Core.dll'
    foreach($generic in @($false,$true)) {
        foreach($outcome in @('success','fault','cancel','completed')) {
            $probe=[UiContinuationProbe]::Run($assembly,$canvas,$outcome,$generic)
            $deadline=[DateTime]::UtcNow.AddSeconds(5)
            while(!$probe.IsCompleted -and [DateTime]::UtcNow -lt $deadline) {
                [Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 5
            }
            if(!$probe.IsCompleted){throw "UI continuation timed out: $outcome"}
            Write-Host $probe.GetAwaiter().GetResult()
        }
    }
    # Thirty-page native PDF: scroll to arbitrary pages with a bounded bitmap cache.
    $pageTotal=30; $fontId=3+2*$pageTotal
    $multiObjects=[Collections.Generic.List[string]]::new()
    $multiObjects.Add('<< /Type /Catalog /Pages 2 0 R >>')
    $kids=(0..($pageTotal-1) | ForEach-Object {"$(3+2*$_) 0 R"}) -join ' '
    $multiObjects.Add("<< /Type /Pages /Kids [$kids] /Count $pageTotal >>")
    for($i=0;$i -lt $pageTotal;$i++) {
        $contentId=4+2*$i
        $multiObjects.Add("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 900] /Resources << /Font << /F1 $fontId 0 R >> >> /Contents $contentId 0 R >>")
        $content="BT /F1 24 Tf 60 820 Td (PAGE $($i+1)) Tj ET"
        $multiObjects.Add("<< /Length $($content.Length) >>`nstream`n$content`nendstream")
    }
    $multiObjects.Add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
    $multiText="%PDF-1.4`n"; $multiOffsets=[Collections.Generic.List[int]]::new()
    for($i=0;$i -lt $multiObjects.Count;$i++){$multiOffsets.Add($multiText.Length);$multiText+="$($i+1) 0 obj`n$($multiObjects[$i])`nendobj`n"}
    $multiXref=$multiText.Length;$multiText+="xref`n0 $($multiObjects.Count+1)`n0000000000 65535 f `n"
    foreach($offset in $multiOffsets){$multiText+=('{0:D10} 00000 n ' -f $offset)+"`n"}
    $multiText+="trailer`n<< /Size $($multiObjects.Count+1) /Root 1 0 R >>`nstartxref`n$multiXref`n%%EOF"
    $multiPath=Join-Path $temp 'continuous.pdf';[IO.File]::WriteAllText($multiPath,$multiText,[Text.Encoding]::ASCII)
    $pdfRegion=$ocr.RecognizeRegion($multiPath,2,[Drawing.RectangleF]::new(.05,.03,.5,.12),$false,[Threading.CancellationToken]::None)
    if($pdfRegion.Text -notmatch 'PAGE' -or $pdfRegion.PageNumber -ne 2){throw 'PDF source crop/OCR in child failed.'}
    Write-Host 'PASS: PDF rasterization and region OCR outside the host process'
    # Non-consecutive page selection, same text/locations, reused engine, and measured latency.
    $watch=[Diagnostics.Stopwatch]::StartNew()
    $batchPages=$ocr.RecognizePages($multiPath,[int[]]@(3,1,2),$null,[Threading.CancellationToken]::None)
    $batchMs=$watch.ElapsedMilliseconds
    if((($batchPages | ForEach-Object PageNumber) -join ',') -ne '3,1,2'){throw 'Batch page identities or order changed.'}
    $watch.Restart()
    for($i=0;$i -lt 3;$i++) {
        $single=$ocr.RecognizeRegion($multiPath,$batchPages[$i].PageNumber,[Drawing.RectangleF]::new(0,0,1,1),$false,[Threading.CancellationToken]::None)
        if($batchPages[$i].Text -ne $single.Text -or $batchPages[$i].Words.Count -eq 0){throw 'Batch lost OCR text or word boxes.'}
    }
    Write-Host ('PASS: selected PDF pages, OCR equivalence; batch_ms='+$batchMs+' separate_ms='+$watch.ElapsedMilliseconds)
    $scanKey=$scanDoc.IndexKey
    $scanScope=[Collections.Generic.List[Doctracker.Core.Models.DocumentRecord]]::new();$scanScope.Add($scanDoc)
    $indexer.IndexMissing($scanState,$null,[Threading.CancellationToken]::None,$scanScope,$true,$false,$true) | Out-Null
    if($scanDoc.IndexKey -ne $scanKey){throw 'Prepared document was recognized a second time.'}
    # Use real isolated child processes to prove concurrency, rollback and cancellation.
    $fakePoolRoot=Join-Path $temp 'pool-worker';[void][IO.Directory]::CreateDirectory($fakePoolRoot)
    $fakePoolWorker=Join-Path $fakePoolRoot 'parallel-worker.exe'
    Add-Type -Path (Join-Path $PSScriptRoot 'ParallelFakeWorker.cs') -OutputAssembly $fakePoolWorker -OutputType ConsoleApplication -ReferencedAssemblies 'System.Xml.dll','System.Core.dll'
    Add-Type -Path (Join-Path $PSScriptRoot 'ParallelOcrProbe.cs') -ReferencedAssemblies (Join-Path $root 'Doctracker.Core.dll'),'System.Drawing.dll','System.Core.dll'
    Write-Host ([ParallelOcrProbe]::Run($assembly,$temp,$multiPath,$fakePoolWorker))
    if(@(Get-ChildItem $workRoot -Directory | Where-Object {$_.Name -notin $beforeWork}).Count -ne 0){throw 'Parallel pool leaked temporary work folders.'}
    # Also run the actual PDFium/Tesseract workers in parallel, including indexed word boxes.
    $parallelDoc=$importer.Import($state,$multiPath,'parallel-real')
    $indexer.WorkerCount=2
    $watch.Restart()
    $indexer.Index($state,$parallelDoc,$null,[Threading.CancellationToken]::None,$true,$true)
    if(!$parallelDoc.IndexComplete -or $parallelDoc.IndexedPages.Count -ne 30){throw 'Real parallel OCR did not index all PDF pages.'}
    for($i=0;$i -lt 30;$i++) {
        $page=$parallelDoc.IndexedPages[$i]
        if($page.PageNumber -ne $i+1 -or $page.Text -notmatch 'PAGE' -or $page.Words.Count -eq 0){throw 'Parallel OCR lost page order, text or word coordinates.'}
    }
    Write-Host ('PASS: actual parallel PDF OCR, 30 pages, two workers; elapsed_ms='+$watch.ElapsedMilliseconds)
    . (Join-Path $PSScriptRoot '../../installer/Ocr-Settings.ps1')
    foreach($spec in @(@(12,16),@(12,8),@(2,16),@(32,64),@(12,0))) {
        $advice=Get-DoctrackerOcrAdvice $spec[0] $spec[1]
        if($advice.Recommended -ne [Doctracker.Core.Services.OcrConcurrencyPolicy]::Recommended($spec[0],$spec[1]) -or $advice.Maximum -ne [Doctracker.Core.Services.OcrConcurrencyPolicy]::Maximum($spec[0])){throw 'Installer and add-in OCR advice differ.'}
    }
    $settingsType=$assembly.GetType('Doctracker.AddIn.Infrastructure.OcrSettings',$true)
    $settingsFlags=[Reflection.BindingFlags]'Public,NonPublic,Static'
    $settingsPath=$settingsType.GetProperty('SettingsPath',$settingsFlags).GetValue($null,$null)
    $previousSettings=if(Test-Path $settingsPath){[IO.File]::ReadAllBytes($settingsPath)}else{$null}
    try {
        # First install, followed by the user's failing case: replace an existing 10.
        if(Test-Path $settingsPath){Remove-Item -LiteralPath $settingsPath}
        Save-DoctrackerOcrWorkers 10
        Save-DoctrackerOcrWorkers 10
        if([IO.File]::ReadAllText($settingsPath) -ne '10'){throw 'Reinstall lost the existing OCR worker setting.'}
        Save-DoctrackerOcrWorkers 6
        if([IO.File]::ReadAllText($settingsPath) -ne '6'){throw 'Reinstall did not update the OCR worker setting.'}
        $locked=[IO.File]::Open($settingsPath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
        $saveRejected=$false
        try {Save-DoctrackerOcrWorkers 4}catch{$saveRejected=$true}finally{$locked.Dispose()}
        if(!$saveRejected -or [IO.File]::ReadAllText($settingsPath) -ne '6'){throw 'Failed settings replacement did not preserve the previous choice.'}
        if(@(Get-ChildItem -LiteralPath (Split-Path $settingsPath) -Filter 'ocr-workers.txt.*.tmp').Count -ne 0){throw 'Settings replacement leaked a temporary file.'}
        Write-Host 'PASS: first install, reinstall with existing OCR setting, changed choice, failed replacement rollback and cleanup'
        Save-DoctrackerOcrWorkers 1
        if($settingsType.GetMethod('LoadWorkers',$settingsFlags).Invoke($null,@()) -ne 1){throw 'Installer worker choice was not loaded by the add-in.'}
        $settingsType.GetMethod('SaveWorkers',$settingsFlags).Invoke($null,[object[]]@(2))
        $expected=[Doctracker.Core.Services.OcrConcurrencyPolicy]::Clamp(2,[Environment]::ProcessorCount)
        if([int][IO.File]::ReadAllText($settingsPath) -ne $expected){throw 'Ribbon worker setting was not persisted.'}
        [IO.File]::WriteAllText($settingsPath,'invalid')
        if($settingsType.GetMethod('LoadWorkers',$settingsFlags).Invoke($null,@()) -ne $settingsType.GetProperty('Recommended',$settingsFlags).GetValue($null,$null)){throw 'Invalid worker settings did not fall back to advice.'}
    } finally {
        if($null -ne $previousSettings){[IO.File]::WriteAllBytes($settingsPath,$previousSettings)}else{Remove-Item $settingsPath -ErrorAction SilentlyContinue}
    }
    Write-Host 'PASS: hardware advice, installer choice, persisted worker count and invalid setting recovery'
    # Verify real checked selection survives changing folder, then clears only visible rows.
    $selectionState=[Doctracker.Core.Models.ProjectState]::new()
    foreach($entry in @(@{Name='BL 01.pdf';Folder='BL'},@{Name='BL 02.pdf';Folder='BL / 2026'},@{Name='Facture 03.pdf';Folder='Factures'})) {
        $doc=[Doctracker.Core.Models.DocumentRecord]::new();$doc.OriginalName=$entry.Name;$doc.Categories.Add($entry.Folder);$selectionState.Documents.Add($doc)
    }
    $pickerType=$assembly.GetType('Doctracker.AddIn.UI.OcrSelectionDialog',$true)
    $picker=$pickerType.GetConstructors($flags)[0].Invoke([object[]]@($selectionState,'BL',$null))
    try {
        $picker.Show();[Windows.Forms.Application]::DoEvents()
        $fileList=$pickerType.GetField('files',$flags).GetValue($picker)
        if($fileList.Items.Count -ne 2){throw 'OCR picker did not restrict to BL and its subfolders.'}
        $pickerType.GetMethod('CheckVisible',$flags).Invoke($picker,[object[]]@($true)) | Out-Null
        $folderList=$pickerType.GetField('folders',$flags).GetValue($picker)
        $folderList.SelectedIndex=0
        if($picker.SelectedDocuments.Count -ne 2){throw 'OCR selection lost on folder change.'}
        $pickerType.GetMethod('CheckVisible',$flags).Invoke($picker,[object[]]@($true)) | Out-Null
        if($picker.SelectedDocuments.Count -ne 3){throw 'OCR multi-selection failed.'}
        $folderList.SelectedItem=@($folderList.Items | Where-Object {$_.Path -eq 'BL'})[0]
        $pickerType.GetMethod('CheckVisible',$flags).Invoke($picker,[object[]]@($false)) | Out-Null
        if($picker.SelectedDocuments.Count -ne 1 -or $picker.SelectedDocuments[0].OriginalName -ne 'Facture 03.pdf'){throw 'Folder deselection affected other folders.'}
        $folderList.SelectedIndex=0
        $bitmap=[Drawing.Bitmap]::new($picker.Width,$picker.Height)
        try {$picker.DrawToBitmap($bitmap,[Drawing.Rectangle]::new(0,0,$picker.Width,$picker.Height));$bitmap.Save((Join-Path $previewDirectory 'ocr-selection.png'))}
        finally {$bitmap.Dispose()}
    } finally {$picker.Close();$picker.Dispose()}
    Write-Host 'PASS: OCR picker multi-selection, folder scope, preserved checks and prepared-index reuse'

    $canvas.LoadDocument($multiPath);$canvas.GoToPage(15)
    if($canvas.CurrentPageNumber -ne 15){throw 'Direct page navigation failed.'}
    $viewport=$type.GetField('viewport',$flags).GetValue($canvas)
    $bounds=$type.GetField('pageBounds',$flags).GetValue($canvas)
    $viewport.AutoScrollPosition=[Drawing.Point]::new(0,$bounds[20].Top)
    $type.GetMethod('RefreshVisiblePages',$flags).Invoke($canvas,@()) | Out-Null
    if($canvas.CurrentPageNumber -ne 21){throw 'Vertical scrolling did not activate page 21.'}
    $surfaces=$type.GetField('pagePictures',$flags).GetValue($canvas)
    if($surfaces.Count -gt 3){throw 'Continuous viewer retained too many rendered pages.'}
    # Zoom beyond the viewport, scroll right, then refresh and switch pages.
    $type.GetMethod('SetZoom',$flags).Invoke($canvas,[object[]]@(1.25,$false)) | Out-Null
    $canvas.GoToPage(21)
    $hostPanel=$type.GetField('pageHost',$flags).GetValue($canvas)
    if(!$viewport.HorizontalScroll.Visible){throw 'Horizontal scrollbar missing after zoom.'}
    $viewport.AutoScrollPosition=[Drawing.Point]::new(350,$bounds[20].Top)
    for($refresh=0;$refresh -lt 5;$refresh++) {
        $type.GetMethod('RefreshVisiblePages',$flags).Invoke($canvas,@()) | Out-Null
        [Windows.Forms.Application]::DoEvents()
    }
    if($viewport.AutoScrollPosition.X -gt -340){throw 'Refreshing zoomed pages reset horizontal scrolling.'}
    $surface=$type.GetField('picture',$flags).GetValue($canvas)
    if($surface.Parent -ne $hostPanel -or $surface.Left -ne $bounds[20].Left){throw 'Page coordinates changed during scroll.'}
    $screenLeft=$viewport.PointToClient($surface.PointToScreen([Drawing.Point]::Empty)).X
    if($screenLeft -gt -300){throw 'Zoomed PDF did not move left on screen.'}
    $canvas.GoToPage(22)
    if($viewport.AutoScrollPosition.X -gt -340){throw 'Page navigation reset horizontal offset.'}
    $viewport.AutoScrollPosition=[Drawing.Point]::new(100000,$bounds[21].Top)
    $type.GetMethod('RefreshVisiblePages',$flags).Invoke($canvas,@()) | Out-Null
    $surface=$type.GetField('picture',$flags).GetValue($canvas)
    $right=$viewport.PointToClient($surface.PointToScreen([Drawing.Point]::new($surface.Width,0))).X
    if($right -gt ($viewport.ClientSize.Width+8)){throw 'Right edge of PDF is unreachable.'}
    $canvas.GoToPage(21)
    Write-Host 'PASS: horizontal scroll at zoom, repeated refresh, page navigation, reachable right edge'
    # Drag an existing snip and cancel it; no new snip is created.
    $editable=New-Object Doctracker.Core.Models.SnipRecord
    $editable.PageNumber=21;$editable.X=.2;$editable.Y=.2;$editable.Width=.3;$editable.Height=.15
    $proofList=[Collections.Generic.List[Doctracker.Core.Models.SnipRecord]]::new();$proofList.Add($editable)
    $canvas.SetProofs($proofList);$canvas.NavigateTo($multiPath,$editable)
    $surface=$type.GetField('picture',$flags).GetValue($canvas)
    $down=[Windows.Forms.MouseEventArgs]::new([Windows.Forms.MouseButtons]::Left,1,[int]($surface.Width*.3),[int]($surface.Height*.25),0)
    $type.GetMethod('Picture_MouseDown',$flags).Invoke($canvas,[object[]]@($surface,$down)) | Out-Null
    if(!$type.GetField('proofDragging',$flags).GetValue($canvas)){throw 'Existing snip did not enter graphical edit mode.'}
    $type.GetMethod('CancelCommentDrag',$flags).Invoke($canvas,@()) | Out-Null
    if($type.GetField('proofDragging',$flags).GetValue($canvas) -or [Math]::Abs($canvas.GetNormalizedSelection().X-.2) -gt .001){throw 'Escape did not restore snip geometry.'}
    $type.GetMethod('SetZoom',$flags).Invoke($canvas,[object[]]@(.2,$false)) | Out-Null
    $canvas.GoToPage(21)
    $continuousShot=[Drawing.Bitmap]::new($canvas.Width,$canvas.Height)
    $canvas.DrawToBitmap($continuousShot,[Drawing.Rectangle]::new(0,0,$canvas.Width,$canvas.Height))
    $continuousShot.Save((Join-Path $previewDirectory 'continuous-pages.png'));$continuousShot.Dispose()
    Write-Host 'PASS: continuous vertical scrolling, direct page number, bounded rendering cache, snip geometry edit/cancel'
    # A new page control can synchronously request another refresh during layout.
    $script:nestedRefreshes=0
    $nestedLayout=[Windows.Forms.ControlEventHandler]{param($sender,$args)
        $script:nestedRefreshes++
        $type.GetMethod('RefreshVisiblePages',$flags).Invoke($canvas,@()) | Out-Null
        $type.GetMethod('UpdatePictureLayout',$flags).Invoke($canvas,@()) | Out-Null
    }
    $hostPanel.add_ControlAdded($nestedLayout)
    try {
        for($repeat=0;$repeat -lt 32;$repeat++) {
            $canvas.Enabled=$false
            $editable.PageNumber=2+($repeat % 27)
            $editable.Width=.2+($repeat % 3)*.03
            $canvas.SetProofs($proofList)
            $canvas.NavigateTo($multiPath,$editable)
            $type.GetMethod('SetZoom',$flags).Invoke($canvas,[object[]]@((.25+($repeat % 3)*.15),$false)) | Out-Null
            $canvas.Enabled=$true
            [Windows.Forms.Application]::DoEvents()
            $activeBefore=$canvas.CurrentPageNumber
            $shot=[Drawing.Bitmap]::new($canvas.Width,$canvas.Height)
            try {$canvas.DrawToBitmap($shot,[Drawing.Rectangle]::new(0,0,$canvas.Width,$canvas.Height))}
            finally {$shot.Dispose()}
            if($canvas.CurrentPageNumber -ne $activeBefore){throw 'Painting changed the active page.'}
            foreach($pageSurface in $surfaces.Values) {
                if($pageSurface.IsDisposed -or $null -eq $pageSurface.Image -or $pageSurface.Image.Width -lt 1){throw 'Preview was retired while still visible.'}
            }
            if($surfaces.Count -gt 4){throw 'Repeated snip navigation leaked cached pages.'}
        }
    }
    finally {$hostPanel.remove_ControlAdded($nestedLayout);$canvas.Enabled=$true}
    if($script:nestedRefreshes -lt 5){throw 'Nested layout regression was not exercised.'}
    Write-Host 'PASS: 32 snip resize/navigation/zoom cycles, nested layout, bitmap lifetime and stable active page'
    # Many visible pages at very low zoom must not retain full-resolution bitmaps.
    $type.GetMethod('SetZoom',$flags).Invoke($canvas,[object[]]@(.02,$false)) | Out-Null
    $canvas.GoToPage(1)
    $surfaces=$type.GetField('pagePictures',$flags).GetValue($canvas)
    [long]$pixelBytes=0
    foreach($entry in $surfaces.Values){$pixelBytes += [long]$entry.Image.Width*$entry.Image.Height*4}
    if($pixelBytes -gt 32MB){throw "Low-zoom preview cache exceeds 32 MB: $pixelBytes"}
    $region=[Drawing.RectangleF]::new(.1,.1,.3,.2)
    for($repeat=0;$repeat -lt 40;$repeat++) {
        $canvas.GoToPage(($repeat % 30)+1)
        $crop=$canvas.CropPageRegion(($repeat % 30)+1,$region)
        try {if($crop.Width -lt 449 -or $crop.Height -lt 449){throw 'OCR crop incorrectly used preview resolution.'}}
        finally {$crop.Dispose()}
    }
    # The captured page remains explicit even if layout changes the visible page.
    $canvas.GoToPage(20)
    $crop=$canvas.CropPageRegion(1,$region);$crop.Dispose()
    for($repeat=0;$repeat -lt 8;$repeat++) {
        $canvas.LoadDocument($multiPath)
        $canvas.GoToPage(15)
        $canvas.ClearDocument()
        [Windows.Forms.Application]::DoEvents()
    }
    $canvas.LoadDocument($multiPath)
    Write-Host 'PASS: low-zoom bitmap budget, 40 source-resolution crops, explicit crop page, repeated open/close'
    $script:displayFailureCount=0
    $failureHandler=[Action[Exception]]{param($failure) $script:displayFailureCount++}
    $canvas.add_DisplayFailed($failureHandler)
    $fail=[Action]{throw [InvalidOperationException]::new('Synthetic display failure')}
    $guard=$type.GetMethod('GuardDisplay',$flags)
    $guard.Invoke($canvas,[object[]]@($fail)) | Out-Null
    $guard.Invoke($canvas,[object[]]@($fail)) | Out-Null
    if($script:displayFailureCount -ne 1){throw 'Display error was not reported once and stopped.'}
    $canvas.LoadDocument($multiPath)
    if($type.GetField('displayFaulted',$flags).GetValue($canvas)){throw 'Reload did not recover the viewer.'}
    $canvas.remove_DisplayFailed($failureHandler)
    Write-Host 'PASS: contained display failure and recovery by reloading'


} finally {
    if ($canvas) { $canvas.Dispose() }
    if ($ocr) { $ocr.Dispose() }
    Remove-Item $temp -Recurse -Force
}
