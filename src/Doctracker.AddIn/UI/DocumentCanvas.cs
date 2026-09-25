using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using Doctracker.AddIn.Infrastructure;
using Doctracker.Core.Models;
using Doctracker.Core.Geometry;
using PdfiumViewer;

namespace Doctracker.AddIn.UI
{
    /// <summary>
    /// PDF/image surface with a real scrollable zoom surface.
    /// Left drag selects a snip; middle drag pans the zoomed document.
    /// </summary>
    internal sealed class DocumentCanvas : UserControl
    {
        private readonly Panel viewport;
        // Only this host moves during scrolling; page bounds stay in document coordinates.
        private readonly Panel pageHost = new Panel {Location=Point.Empty,Margin=Padding.Empty};
        private PictureBox picture;
        private readonly PictureBox emptyPicture;
        private readonly Dictionary<int,PictureBox> pagePictures = new Dictionary<int,PictureBox>();
        private readonly Dictionary<int,Size> previewSizes = new Dictionary<int,Size>();
        private bool changingPreview, refreshQueued, layoutQueued;
        private readonly List<Rectangle> pageBounds = new List<Rectangle>();
        private bool layingOutPages, refreshingPages;
        private readonly NumericUpDown pageNumber = new NumericUpDown {Minimum=1,Maximum=1,Width=58,Value=1};
        private bool updatingPageNumber;
        private readonly Label pageLabel;
        private readonly Label zoomLabel;
        private readonly Button previousButton;
        private readonly Button nextButton;
        private readonly Button zoomOutButton;
        private readonly Button fitButton;
        private readonly Button zoomInButton;

        private PdfDocument pdf;
        private Image currentImage => currentPath==null ? null : picture?.Image;
        private Size imageNaturalSize;
        private bool displayFaulted;
        public event Action<Exception> DisplayFailed;
        private string currentPath;
        private int pageIndex;
        private int imagePageCount = 1;
        private double zoom = 1d;
        private bool fitToViewport = true;
        private bool fitWidth = true;
        private DocumentRecord document;
        private readonly ContextMenuStrip popup = new ContextMenuStrip();
        private readonly ContextMenuStrip zoomMenu = new ContextMenuStrip();
        public bool CommentMode { get; set; }
        private string selectedCommentId;
        private DocumentComment commentPreview;
        private RectangleF commentOriginal;
        private bool commentDragging;
        private bool proofDragging;
        private RectangleF proofOriginal, proofPreview;
        private int proofHandle;
        public event Action<string,RectangleF> ProofGeometryChanged;
        private int commentHandle;
        public event Action<string,RectangleF> CommentGeometryChanged;
        public event Action<string> DeleteProofRequested;
        public event Action<string> ProofSelected;
        public event Action<string> EditCommentRequested;
        public event Action<string> DeleteCommentRequested;
        public void SetDocument(DocumentRecord value) { if(document?.Id!=value?.Id)selectedCommentId=null;document = value; InvalidatePages(); }
        private Point dragStart;
        private Point dragEnd;
        private bool dragging;
        private bool panning;
        private Point panStart;
        private Point panOrigin;
        private RectangleF? normalizedSelection;
        private readonly List<SnipRecord> proofs = new List<SnipRecord>();
        private SnipType selectionType = SnipType.Text;
        private string selectedProofId;
        public SnipType? ActiveType { get; set; } = SnipType.Text;
        public void SetProofs(IEnumerable<SnipRecord> records)
        { proofs.Clear(); proofs.AddRange(records); InvalidatePages(); }


        public DocumentCanvas()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);

            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96,96);
            BackColor = SnipTheme.Surface;

            var toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(6, 5, 6, 4),
                BackColor = Color.White
            };

            previousButton = CreateToolbarButton("", "Page précédente", "Previous");
            nextButton = CreateToolbarButton("", "Page suivante", "Next");
            pageLabel = new Label
            {
                Text = "Aucun document",
                AutoSize = true,
                Padding = new Padding(8, 5, 8, 0),
                ForeColor = Color.FromArgb(45, 53, 64)
            };
            zoomOutButton = CreateToolbarButton("", "Réduire le zoom", "Minus");
            fitButton = CreateToolbarButton("Largeur", "Adapter à la largeur ; clic droit pour voir la page entière", "Fit");
            zoomInButton = CreateToolbarButton("", "Agrandir le zoom", "Plus");
            zoomLabel = new Label
            {
                Text = "100 %",
                AutoSize = true,
                Padding = new Padding(4, 5, 4, 0),
                ForeColor = Color.FromArgb(80, 88, 98)
            };

            previousButton.Click += (sender, args) => ShowPage(pageIndex - 1);
            nextButton.Click += (sender, args) => ShowPage(pageIndex + 1);
            zoomOutButton.Click += (sender, args) => SetZoom(zoom - 0.15d, false);
            fitButton.Click += (sender, args) => FitWidth();
            zoomMenu.Items.Add("Adapter à la largeur", null, (s,e) => FitWidth());
            zoomMenu.Items.Add("Page entière", null, (s,e) => FitPage());
            zoomMenu.Items.Add("100 %", null, (s,e) => SetZoom(1d / DisplayScale(), false));
            fitButton.ContextMenuStrip = zoomMenu;
            zoomInButton.Click += (sender, args) => SetZoom(zoom + 0.15d, false);

            toolbar.Controls.Add(previousButton);
            toolbar.Controls.Add(nextButton);
            pageNumber.AccessibleName="Aller à la page";
            pageNumber.ValueChanged+=(s,e)=>{if(!updatingPageNumber)ShowPage((int)pageNumber.Value-1);};
            toolbar.Controls.Add(pageNumber);
            toolbar.Controls.Add(pageLabel);
            toolbar.Controls.Add(zoomOutButton);
            toolbar.Controls.Add(fitButton);
            toolbar.Controls.Add(zoomInButton);
            toolbar.Controls.Add(zoomLabel);

            viewport = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.FromArgb(235, 239, 245),
                Padding = new Padding(4)
            };
            viewport.Resize += (sender, args) =>
            {
                if (currentImage != null) UpdatePictureLayout();
            };

            picture = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.StretchImage,
                BackColor = Color.White,
                Cursor = Cursors.Cross,
                TabStop = false
            };
            emptyPicture=picture; picture.Tag=0;
            picture.MouseDown += Picture_MouseDown;
            picture.MouseMove += Picture_MouseMove;
            picture.MouseUp += Picture_MouseUp;
            picture.MouseDoubleClick+=(s,e)=>{
                if(e.Button!=MouseButtons.Left)return;
                var comment=CommentAt(e.Location);if(comment==null)return;
                commentDragging=false;commentPreview=null;picture.Capture=false;EditCommentRequested?.Invoke(comment.Id);picture.Invalidate();
            };
            picture.KeyDown+=(s,e)=>{if(e.KeyCode==Keys.Escape)CancelCommentDrag();};
            viewport.KeyDown+=(s,e)=>{if(e.KeyCode==Keys.Escape)CancelCommentDrag();};
            picture.MouseCaptureChanged+=(s,e)=>{if(!picture.Capture && (commentDragging || proofDragging))CancelCommentDrag();};
            picture.Paint += PaintPage;
            // Hover must not steal keyboard focus from Excel or the search box.
            MouseEventHandler wheel = (s,e) => {
                if ((ModifierKeys & Keys.Control) == 0)
                { if(IsHandleCreated)BeginInvoke(new Action(RefreshVisiblePages)); return; }
                SetZoom(zoom * (e.Delta > 0 ? 1.15 : 1 / 1.15), false);
                if (e is HandledMouseEventArgs handled) handled.Handled = true;
            };
            picture.MouseWheel += wheel; viewport.MouseWheel += wheel;

            viewport.Scroll+=(s,e)=>RefreshVisiblePages();
            pageHost.Controls.Add(picture);
            viewport.Controls.Add(pageHost);
            Controls.Add(viewport);
            Controls.Add(toolbar);

            UpdateNavigationState();
        }

        public event EventHandler SelectionCompleted;

        public string CurrentPath => currentPath;
        public void ClearDocument() => DisposeDocument();
        public int CurrentPageNumber => pageIndex + 1;
        public bool HasDocument => currentImage != null;
        public bool HasSelection => normalizedSelection.HasValue;

        public void LoadDocument(string path)
        {
            DisposeDocument();displayFaulted=false;
            proofs.Clear(); selectedProofId = null;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                throw new FileNotFoundException("The document cannot be found.", path);
            }

            try
            {
                currentPath = path;
                pageIndex = 0;
                fitToViewport = true; fitWidth = true;
                normalizedSelection = null;

                if (string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    NativePdfiumLoader.EnsureLoaded();
                    pdf = PdfDocument.Load(path);
                }
                if (pdf == null)
                    using (var image = Image.FromFile(path)) {imagePageCount = DocumentIndexer.ImagePageCount(image);imageNaturalSize=image.Size;}
                RenderCurrentPage();
            }
            catch
            {
                DisposeDocument();
                throw;
            }
        }

        public void NavigateTo(string path, SnipRecord snip)
        {
            if (!string.Equals(currentPath, path, StringComparison.OrdinalIgnoreCase))
            {
                LoadDocument(path);
            }

            ShowPage(Math.Max(0, snip.PageNumber - 1));
            selectionType = snip.SourceType ?? snip.Type;
            selectedProofId = snip.Id;
            normalizedSelection = new RectangleF(
                (float)snip.X, (float)snip.Y, (float)snip.Width, (float)snip.Height);
            RevealSelection();
            picture.Invalidate();
        }

        public void NavigateTo(string path, int targetPageNumber)
        {
            if (!string.Equals(currentPath, path, StringComparison.OrdinalIgnoreCase))
            {
                LoadDocument(path);
            }

            ShowPage(Math.Max(0, targetPageNumber - 1));
            normalizedSelection = null;
            FitWidth();
            picture.Invalidate();
        }

        public RectangleF GetNormalizedSelection()
        {
            if (!normalizedSelection.HasValue)
                throw new InvalidOperationException("Draw a zone on the document first.");
            return normalizedSelection.Value;
        }

        public Bitmap CropSelection() => CropRegion(GetNormalizedSelection());
        public Bitmap CropRegion(RectangleF selection) => CropPageRegion(CurrentPageNumber,selection);
        public Bitmap CropPageRegion(int number,RectangleF selection)
        {
            var sourceIndex=number-1;
            if(sourceIndex<0 || sourceIndex>=PageCount)throw new InvalidOperationException("La page sélectionnée n'est plus disponible.");
            if(currentImage==null)throw new InvalidOperationException("Aucun document ouvert.");
            // OCR always reads source pixels, independently of the lightweight preview.
            if(pdf!=null)
            {
                var size=DocumentIndexer.RenderSize(pdf.PageSizes[sourceIndex]);
                using(var source=pdf.Render(sourceIndex,size.Width,size.Height,144,144,PdfRenderFlags.Annotations))
                    return CropImage(source,selection);
            }
            using(var source=Image.FromFile(currentPath))
            {
                if(imagePageCount>1)source.SelectActiveFrame(FrameDimension.Page,sourceIndex);
                return CropImage(source,selection);
            }
        }
        private static Bitmap CropImage(Image source,RectangleF selection)
        {
            var crop=Rectangle.FromLTRB(Math.Max(0,(int)Math.Floor(selection.Left*source.Width)),Math.Max(0,(int)Math.Floor(selection.Top*source.Height)),
                Math.Min(source.Width,(int)Math.Ceiling(selection.Right*source.Width)),Math.Min(source.Height,(int)Math.Ceiling(selection.Bottom*source.Height)));
            if(crop.Width<2 || crop.Height<2)throw new InvalidOperationException("La zone sélectionnée est trop petite.");
            var scale=Math.Min(1,3000d/Math.Max(crop.Width,crop.Height));
            var output=new Bitmap(Math.Max(1,(int)Math.Ceiling(crop.Width*scale)),Math.Max(1,(int)Math.Ceiling(crop.Height*scale)));
            try
            {
                using(var graphics=Graphics.FromImage(output))graphics.DrawImage(source,new Rectangle(Point.Empty,output.Size),crop,GraphicsUnit.Pixel);
                return output;
            }
            catch {output.Dispose();throw;}
        }

        public void ClearSelection()
        {
            normalizedSelection = null;
            selectedProofId = null;
            picture.Invalidate();
        }

        private int PageCount => pdf==null ? (currentPath==null?0:imagePageCount) : pdf.PageCount;
        private void InvalidatePages() { foreach(var surface in pagePictures.Values)surface.Invalidate();picture.Invalidate(); }
        public void GoToPage(int number) => ShowPage(number-1);
        private void ShowPage(int requestedIndex) => GuardDisplay(()=>ShowPageCore(requestedIndex));
        private void ShowPageCore(int requestedIndex)
        {
            if(requestedIndex<0 || requestedIndex>=PageCount)return;
            CancelCommentDrag(); selectedCommentId=null; selectedProofId=null; normalizedSelection=null;
            ActivatePage(GetPagePicture(requestedIndex));
            if(pageBounds.Count!=PageCount)UpdatePictureLayout();
            viewport.AutoScrollPosition=new Point(Math.Max(0,-viewport.AutoScrollPosition.X),Math.Max(0,pageBounds[requestedIndex].Top-4));
            RefreshVisiblePages(); UpdateNavigationState();
        }
        private void ActivatePage(PictureBox surface)
        {
            if(surface==null || surface.Image==null)return;
            var next=(int)surface.Tag;
            if(pageIndex!=next){normalizedSelection=null;selectedProofId=null;selectedCommentId=null;}
            picture=surface;pageIndex=next;UpdateNavigationState();
        }
        private PictureBox GetPagePicture(int index)
        {
            // Image.Size calls GDI+; bookkeeping must not inspect a bitmap while it is painted.
            // Image/control assignment can trigger nested layout synchronously.
            changingPreview=true;
            try { return GetPagePictureCore(index); }
            finally { changingPreview=false; }
        }
        private PictureBox GetPagePictureCore(int index)
        {
            var size=PreviewSize(index);
            if(pagePictures.TryGetValue(index,out var existing))
            {
                if(!previewSizes.TryGetValue(index,out var renderedSize) || renderedSize!=size)
                {
                    var replacement=RenderPreview(index,size);var old=existing.Image;
                    existing.Image=replacement;previewSizes[index]=size;
                    old?.Dispose();
                }
                return existing;
            }
            var surface=index==0?emptyPicture:new PictureBox {SizeMode=PictureBoxSizeMode.StretchImage,BackColor=Color.White,Cursor=Cursors.Cross,TabStop=false};
            surface.Tag=index;
            if(surface!=emptyPicture)
            {
                surface.MouseDown+=Picture_MouseDown;surface.MouseMove+=Picture_MouseMove;surface.MouseUp+=Picture_MouseUp;surface.Paint+=PaintPage;
                surface.MouseDoubleClick+=(s,e)=>{ActivatePage(surface);var comment=CommentAt(e.Location);if(e.Button==MouseButtons.Left && comment!=null){CancelCommentDrag();EditCommentRequested?.Invoke(comment.Id);}};
                surface.MouseCaptureChanged+=(s,e)=>{if(!surface.Capture && (commentDragging || proofDragging))CancelCommentDrag();};
                surface.MouseWheel+=(s,e)=>{
                    if((ModifierKeys & Keys.Control)!=0){SetZoom(zoom*(e.Delta>0?1.15:1/1.15),false);if(e is HandledMouseEventArgs handled)handled.Handled=true;}
                    else if(IsHandleCreated)BeginInvoke(new Action(RefreshVisiblePages));
                };
                pageHost.Controls.Add(surface);
            }
            try
            {
                surface.Image=RenderPreview(index,size);
                pagePictures[index]=surface;previewSizes[index]=size;return surface;
            }
            catch {if(surface!=emptyPicture){pageHost.Controls.Remove(surface);surface.Dispose();}throw;}
        }
        private Size PreviewSize(int index)
        {
            var natural=pdf==null?imageNaturalSize:DocumentIndexer.RenderSize(pdf.PageSizes[index]);
            var displayed=pageBounds.Count>index?pageBounds[index].Size:new Size(Math.Max(120,viewport.ClientSize.Width),Math.Max(120,viewport.ClientSize.Height));
            var scale=Math.Min(1,Math.Max(128,Math.Max(displayed.Width,displayed.Height)*1.25)/Math.Max(natural.Width,natural.Height));
            scale=Math.Min(scale,3000d/Math.Max(natural.Width,natural.Height));
            return new Size(Math.Max(1,(int)Math.Ceiling(natural.Width*scale)),Math.Max(1,(int)Math.Ceiling(natural.Height*scale)));
        }
        private Image RenderPreview(int index,Size size)
        {
            if(pdf!=null)return pdf.Render(index,size.Width,size.Height,144,144,PdfRenderFlags.Annotations|PdfRenderFlags.LcdText);
            using(var source=Image.FromFile(currentPath))
            {
                if(imagePageCount>1)source.SelectActiveFrame(FrameDimension.Page,index);
                return new Bitmap(source,size);
            }
        }
        private void GuardDisplay(Action action)
        {
            if(IsDisposed || Disposing || displayFaulted)return;
            try {action();}
            catch(Exception failure)
            {
                DiagnosticLog.Write("DisplayFailure",failure);
                if(DisplayFailed==null)throw; // Tests and standalone callers still observe errors.
                displayFaulted=true;dragging=proofDragging=commentDragging=panning=false;
                DisplayFailed(failure);
            }
        }

        private void RenderCurrentPage()
        {
            ActivatePage(GetPagePicture(pageIndex));UpdatePictureLayout();
        }
        private void UpdatePictureLayout() => GuardDisplay(()=>UpdatePictureLayoutCore());
        private void UpdatePictureLayoutCore()
        {
            if(changingPreview){QueuePageRefresh(true);return;}
            if(currentImage==null || layingOutPages)return;
            layingOutPages=true;
            try
            {
                var oldTop=-viewport.AutoScrollPosition.Y;
                var fraction=pageBounds.Count>pageIndex ? (oldTop-pageBounds[pageIndex].Top)/(double)Math.Max(1,pageBounds[pageIndex].Height) : 0;
                var availableWidth=Math.Max(120,viewport.Width-viewport.Padding.Horizontal-SystemInformation.VerticalScrollBarWidth-2);
                var availableHeight=Math.Max(120,viewport.ClientSize.Height-viewport.Padding.Vertical-4);
                pageBounds.Clear();var top=viewport.Padding.Top;var widest=0;
                for(var index=0;index<PageCount;index++)
                {
                    var natural=pdf==null?imageNaturalSize:DocumentIndexer.RenderSize(pdf.PageSizes[index]);
                    var ratio=fitToViewport ? (fitWidth?availableWidth/(double)natural.Width:Math.Min(availableWidth/(double)natural.Width,availableHeight/(double)natural.Height)) : zoom;
                    ratio=Math.Max(.02,Math.Min(3,ratio));if(index==pageIndex)zoom=ratio;
                    var width=Math.Max(2,(int)Math.Round(natural.Width*ratio));var height=Math.Max(2,(int)Math.Round(natural.Height*ratio));
                    pageBounds.Add(new Rectangle(Math.Max(4,(availableWidth-width)/2+4),top,width,height));
                    top=checked(top+height+16);widest=Math.Max(widest,width);
                }
                viewport.AutoScroll=true;
                pageHost.Size=new Size(Math.Max(viewport.ClientSize.Width-8,widest+8),top);
                viewport.AutoScrollMinSize=pageHost.Size;
                viewport.AutoScrollPosition=new Point(Math.Max(0,-viewport.AutoScrollPosition.X),Math.Max(0,pageBounds[pageIndex].Top+(int)(fraction*pageBounds[pageIndex].Height)));
                zoomLabel.Text=Math.Round(zoom*DisplayScale()*100)+" %";
            }
            finally {layingOutPages=false;}
            RefreshVisiblePages();
        }
        private void RefreshVisiblePages() => GuardDisplay(()=>RefreshVisiblePagesCore());
        private void RefreshVisiblePagesCore()
        {
            if(changingPreview){QueuePageRefresh(false);return;}
            if(layingOutPages || refreshingPages || currentPath==null || pageBounds.Count==0 || IsDisposed)return;
            refreshingPages=true;
            try
            {
                var top=-viewport.AutoScrollPosition.Y;var bottom=top+viewport.ClientSize.Height;
                var visible=Enumerable.Range(0,pageBounds.Count).Where(i=>pageBounds[i].Bottom>=top && pageBounds[i].Top<=bottom).ToList();
                if(visible.Count==0)return;
                var active=visible.FirstOrDefault(i=>pageBounds[i].Bottom>top+Math.Min(40,viewport.Height/4));
                if(!dragging && !commentDragging && !proofDragging && !panning)ActivatePage(GetPagePicture(active));
                var keep=new HashSet<int>(visible);keep.Add(pageIndex);
                foreach(var old in pagePictures.Keys.Where(i=>!keep.Contains(i)).ToList())
                {
                    var surface=pagePictures[old];pagePictures.Remove(old);previewSizes.Remove(old);var image=surface.Image;surface.Image=null;image?.Dispose();
                    if(surface!=emptyPicture){pageHost.Controls.Remove(surface);surface.Dispose();}else surface.Visible=false;
                }
                foreach(var index in keep)
                {
                    var surface=GetPagePicture(index);var bounds=pageBounds[index];
                    surface.Bounds=bounds;surface.Visible=true;surface.Invalidate();
                }
            }
            finally {refreshingPages=false;}
        }
        private void QueuePageRefresh(bool layout)
        {
            layoutQueued |= layout;
            if(refreshQueued || !IsHandleCreated || IsDisposed)return;
            refreshQueued=true;
            BeginInvoke(new Action(()=>{
                var needsLayout=layoutQueued;layoutQueued=false;refreshQueued=false;
                if(needsLayout)UpdatePictureLayout();else RefreshVisiblePages();
            }));
        }
        private void PaintPage(object sender,PaintEventArgs e) => GuardDisplay(()=>Picture_Paint(sender,e));

        private void SetZoom(double requestedZoom, bool fit) => GuardDisplay(()=>SetZoomCore(requestedZoom,fit));
        private void SetZoomCore(double requestedZoom, bool fit)
        {
            if (currentImage == null) return;
            fitToViewport = fit;
            if (!fitToViewport)
            {
                zoom = Math.Max(0.02d, Math.Min(3.0d, requestedZoom));
            }
            UpdatePictureLayout();
        }

        private double DisplayScale()
        {
            if (currentImage == null) return 1;
            var logicalWidth = pdf != null ? pdf.PageSizes[pageIndex].Width * 96d / 72d : imageNaturalSize.Width;
            return (pdf!=null?DocumentIndexer.RenderSize(pdf.PageSizes[pageIndex]).Width:imageNaturalSize.Width) / Math.Max(1, logicalWidth);
        }
        public void FitWidth() { fitWidth = true; SetZoom(1d, true); }
        public void FitPage() { fitWidth = false; SetZoom(1d, true); }
        public RectangleF FitComment(RectangleF zone, string text, double fontSize = 16)
        {
            // Measure at a stable page width so wrapping is independent of current zoom.
            var natural=pdf==null?imageNaturalSize:DocumentIndexer.RenderSize(pdf.PageSizes[pageIndex]);
            var referenceSize = new Size(1000, (int)(1000d * natural.Height / natural.Width));
            var padding = 12f;
            var width = zone.Width * referenceSize.Width - padding;
            if (width < 40) throw new InvalidOperationException("Dessinez une zone de commentaire plus large.");
            using (var graphics = picture.CreateGraphics())
            using (var font = new Font("Segoe UI", DocumentOverlay.FontPixels(referenceSize.Width, fontSize), FontStyle.Regular, GraphicsUnit.Pixel))
            {
                var needed = (graphics.MeasureString(text, font, (int)width).Height + padding + 6) / referenceSize.Height;
                zone.Height = Math.Max(zone.Height, needed);
                if (zone.Bottom > 1) throw new InvalidOperationException("Ce texte dépasse le bas de la page. Dessinez une zone plus large ou plus haute dans la page.");
                return zone;
            }
        }


        private void Picture_MouseDown(object sender, MouseEventArgs e) => GuardDisplay(()=>Picture_MouseDownCore(sender,e));
        private void Picture_MouseDownCore(object sender, MouseEventArgs e)
        {
            ActivatePage(sender as PictureBox);
            if (currentImage == null) return;
            // Give Escape/zoom keys to the document only after an intentional click.
            viewport.Focus();

            if (e.Button == MouseButtons.Middle)
            {
                panning = true;
                picture.Capture = true;
                panStart = picture.PointToScreen(e.Location);
                panOrigin = viewport.AutoScrollPosition;
                picture.Cursor = Cursors.SizeAll;
                return;
            }

            if (e.Button == MouseButtons.Right)
            {
                popup.Items.Clear();
                var point = new PointF(e.X / (float)picture.Width, e.Y / (float)picture.Height);
                var comment = document?.Comments.LastOrDefault(c => c.PageNumber == CurrentPageNumber && new RectangleF((float)c.X,(float)c.Y,(float)c.Width,(float)c.Height).Contains(point));
                var proof = ProofAt(point);
                if (comment != null)
                {
                    popup.Items.Add("Modifier le commentaire…", null, (s,a) => EditCommentRequested?.Invoke(comment.Id));
                    popup.Items.Add("Supprimer le commentaire", null, (s,a) => DeleteCommentRequested?.Invoke(comment.Id));
                }
                else if (proof != null) popup.Items.Add("Supprimer ce snip", null, (s,a) => DeleteProofRequested?.Invoke(proof.Id));
                if (popup.Items.Count > 0) popup.Show(picture, e.Location);
                return;
            }
            if (e.Button != MouseButtons.Left) return;
            var selected=document?.Comments.FirstOrDefault(c=>c.Id==selectedCommentId && c.PageNumber==CurrentPageNumber);
            var handle=selected==null?-1:CommentHandleAt(selected,e.Location);
            var hit=handle>0?selected:CommentAt(e.Location);
            if(hit!=null)
            {
                selectedCommentId=hit.Id;commentOriginal=new RectangleF((float)hit.X,(float)hit.Y,(float)hit.Width,(float)hit.Height);
                commentHandle=handle>0?handle:0;dragStart=e.Location;
                commentPreview=new DocumentComment {Id=hit.Id,PageNumber=hit.PageNumber,X=hit.X,Y=hit.Y,Width=hit.Width,Height=hit.Height,Text=hit.Text,FontSize=hit.FontSize};
                commentDragging=true;normalizedSelection=null;picture.Capture=true;picture.Invalidate();return;
            }
            selectedCommentId=null;
            var selectedProof=proofs.FirstOrDefault(p=>p.Id==selectedProofId && p.PageNumber==CurrentPageNumber);
            var proofGrip=selectedProof==null?-1:CommentHandleAt(ProofFrame(selectedProof),e.Location);
            var proofHit=proofGrip>0?selectedProof:ProofAt(new PointF(e.X/(float)picture.Width,e.Y/(float)picture.Height));
            if(proofHit!=null && (ModifierKeys & Keys.Alt)==0 && !CommentMode)
            {
                selectedProofId=proofHit.Id;selectionType=proofHit.SourceType??proofHit.Type;
                proofOriginal=new RectangleF((float)proofHit.X,(float)proofHit.Y,(float)proofHit.Width,(float)proofHit.Height);
                proofPreview=proofOriginal;normalizedSelection=proofOriginal;proofHandle=proofGrip>0?proofGrip:0;
                proofDragging=true;dragStart=e.Location;picture.Capture=true;picture.Invalidate();return;
            }
            picture.Capture = true;
            dragging = true;
            selectionType = ActiveType ?? SnipType.Text;
            selectedProofId = null;
            dragStart = Clamp(e.Location, picture.ClientRectangle);
            dragEnd = dragStart;
            normalizedSelection = null;
        }

        private void Picture_MouseMove(object sender, MouseEventArgs e) => GuardDisplay(()=>Picture_MouseMoveCore(sender,e));
        private void Picture_MouseMoveCore(object sender, MouseEventArgs e)
        {
            if(proofDragging)
            {
                var box=CommentGeometry.Transform(new NormalizedRectangle(proofOriginal.X,proofOriginal.Y,proofOriginal.Width,proofOriginal.Height),
                    (e.X-dragStart.X)/(double)picture.Width,(e.Y-dragStart.Y)/(double)picture.Height,proofHandle);
                proofPreview=new RectangleF((float)box.X,(float)box.Y,(float)box.Width,(float)box.Height);normalizedSelection=proofPreview;picture.Invalidate();return;
            }
            if(commentDragging && commentPreview!=null)
            {
                var box=CommentGeometry.Transform(new NormalizedRectangle(commentOriginal.X,commentOriginal.Y,commentOriginal.Width,commentOriginal.Height),
                    (e.X-dragStart.X)/(double)picture.Width,(e.Y-dragStart.Y)/(double)picture.Height,commentHandle);
                commentPreview.X=box.X;commentPreview.Y=box.Y;commentPreview.Width=box.Width;commentPreview.Height=box.Height;picture.Invalidate();return;
            }
            if(!dragging && !panning)
            {
                var selected=document?.Comments.FirstOrDefault(c=>c.Id==selectedCommentId && c.PageNumber==CurrentPageNumber);
                var chosenProof=proofs.FirstOrDefault(p=>p.Id==selectedProofId && p.PageNumber==CurrentPageNumber);
                var handle=selected!=null?CommentHandleAt(selected,e.Location):chosenProof==null?-1:CommentHandleAt(ProofFrame(chosenProof),e.Location);
                picture.Cursor=handle==1||handle==5?Cursors.SizeNWSE:handle==3||handle==7?Cursors.SizeNESW:handle==2||handle==6?Cursors.SizeNS:handle==4||handle==8?Cursors.SizeWE:CommentAt(e.Location)!=null?Cursors.SizeAll:Cursors.Cross;
            }
            if (panning)
            {
                var screen = picture.PointToScreen(e.Location);
                var x = -panOrigin.X - (screen.X - panStart.X);
                var y = -panOrigin.Y - (screen.Y - panStart.Y);
                viewport.AutoScrollPosition = new Point(Math.Max(0, x), Math.Max(0, y));
                return;
            }

            if (!dragging) return;
            dragEnd = Clamp(e.Location, picture.ClientRectangle);
            picture.Invalidate();
        }

        private void Picture_MouseUp(object sender, MouseEventArgs e) => GuardDisplay(()=>Picture_MouseUpCore(sender,e));
        private void Picture_MouseUpCore(object sender, MouseEventArgs e)
        {
            if(proofDragging && e.Button==MouseButtons.Left)
            {
                var id=selectedProofId;var preview=proofPreview;proofDragging=false;picture.Capture=false;
                if(Math.Abs(e.X-dragStart.X)>2 || Math.Abs(e.Y-dragStart.Y)>2)ProofGeometryChanged?.Invoke(id,preview);
                else ProofSelected?.Invoke(id);
                picture.Invalidate();return;
            }
            if(commentDragging && e.Button==MouseButtons.Left)
            {
                var preview=commentPreview;commentDragging=false;commentPreview=null;picture.Capture=false;
                if(preview!=null && (Math.Abs(e.X-dragStart.X)>2 || Math.Abs(e.Y-dragStart.Y)>2))
                    CommentGeometryChanged?.Invoke(preview.Id,new RectangleF((float)preview.X,(float)preview.Y,(float)preview.Width,(float)preview.Height));
                picture.Invalidate();return;
            }
            if (panning && e.Button == MouseButtons.Middle)
            {
                panning = false;
                picture.Capture = false;
                picture.Cursor = Cursors.Cross;
                return;
            }

            if (!dragging || e.Button != MouseButtons.Left) return;
            dragging = false;
            picture.Capture = false;
            dragEnd = Clamp(e.Location, picture.ClientRectangle);
            var rectangle = NormalizeScreenRectangle(dragStart, dragEnd);
            if (rectangle.Width >= 4 && rectangle.Height >= 4)
            {
                normalizedSelection = new RectangleF(
                    rectangle.Left / (float)picture.ClientSize.Width,
                    rectangle.Top / (float)picture.ClientSize.Height,
                    rectangle.Width / (float)picture.ClientSize.Width,
                    rectangle.Height / (float)picture.ClientSize.Height);
                SelectionCompleted?.Invoke(this, EventArgs.Empty);
            }
            else
            {
                var proof = ProofAt(new PointF(e.X / (float)picture.Width, e.Y / (float)picture.Height));
                if (proof != null) { NavigateTo(currentPath, proof); ProofSelected?.Invoke(proof.Id); }
            }
            picture.Invalidate();
        }
        private SnipRecord ProofAt(PointF point) => proofs.LastOrDefault(p => p.PageNumber == CurrentPageNumber &&
            new RectangleF((float)p.X,(float)p.Y,(float)p.Width,(float)p.Height).Contains(point));

        private void Picture_Paint(object sender, PaintEventArgs e)
        {
            var surface=(PictureBox)sender;
            var number=(int)surface.Tag+1;
            var active=ReferenceEquals(surface,picture);
            var preview=active?commentPreview:null;
            DocumentOverlay.Draw(e.Graphics, surface.Size, document, number,preview);
            var selectedComment=preview??document?.Comments.FirstOrDefault(c=>c.Id==selectedCommentId && c.PageNumber==number);
            if(selectedComment!=null)
                foreach(var point in CommentHandles(selectedComment,surface.Size))
                {var size=Math.Max(6,Font.Height/2);var box=new RectangleF(point.X-size/2f,point.Y-size/2f,size,size);e.Graphics.FillRectangle(Brushes.White,box);e.Graphics.DrawRectangle(Pens.Red,box.X,box.Y,box.Width,box.Height);}
            foreach (var proof in proofs.Where(item => item.PageNumber == number && item.Id != selectedProofId))
            {
                var bounds = new Rectangle((int)(proof.X * surface.Width), (int)(proof.Y * surface.Height),
                    (int)(proof.Width * surface.Width), (int)(proof.Height * surface.Height));
                DrawHighlight(e.Graphics, bounds, SnipTheme.ColorFor(proof.SourceType ?? proof.Type), false);
            }
            var selectedSnip=proofs.FirstOrDefault(p=>p.Id==selectedProofId && p.PageNumber==number);
            if(selectedSnip!=null)
            {
                var frame=ProofFrame(selectedSnip);
                if(active && proofDragging){frame.X=proofPreview.X;frame.Y=proofPreview.Y;frame.Width=proofPreview.Width;frame.Height=proofPreview.Height;}
                foreach(var point in CommentHandles(frame,surface.Size))
                {var size=Math.Max(6,Font.Height/2);var box=new RectangleF(point.X-size/2f,point.Y-size/2f,size,size);e.Graphics.FillRectangle(Brushes.White,box);using(var pen=new Pen(SnipTheme.ColorFor(selectedSnip.SourceType??selectedSnip.Type)))e.Graphics.DrawRectangle(pen,box.X,box.Y,box.Width,box.Height);}
            }
            if(!active)return;
            Rectangle rectangle;
            if (dragging)
            {
                rectangle = NormalizeScreenRectangle(dragStart, dragEnd);
            }
            else if (normalizedSelection.HasValue)
            {
                var selection = normalizedSelection.Value;
                rectangle = new Rectangle(
                    (int)(selection.X * surface.ClientSize.Width),
                    (int)(selection.Y * surface.ClientSize.Height),
                    (int)(selection.Width * surface.ClientSize.Width),
                    (int)(selection.Height * surface.ClientSize.Height));
            }
            else
            {
                return;
            }

            DrawHighlight(e.Graphics, rectangle, CommentMode ? Color.Red : SnipTheme.ColorFor(selectionType), true);
        }

        private static DocumentComment ProofFrame(SnipRecord snip) => new DocumentComment {X=snip.X,Y=snip.Y,Width=snip.Width,Height=snip.Height};
        private DocumentComment CommentAt(Point point)=>document?.Comments.LastOrDefault(c=>c.PageNumber==CurrentPageNumber && DocumentOverlay.Bounds(c,picture.Size).Contains(point));
        private PointF[] CommentHandles(DocumentComment comment)
            => CommentHandles(comment,picture.Size);
        private static PointF[] CommentHandles(DocumentComment comment,Size size)
        {
            var r=DocumentOverlay.Bounds(comment,size);var cx=(r.Left+r.Right)/2;var cy=(r.Top+r.Bottom)/2;
            return new[]{new PointF(r.Left,r.Top),new PointF(cx,r.Top),new PointF(r.Right,r.Top),new PointF(r.Right,cy),new PointF(r.Right,r.Bottom),new PointF(cx,r.Bottom),new PointF(r.Left,r.Bottom),new PointF(r.Left,cy)};
        }
        private int CommentHandleAt(DocumentComment comment,Point point)
        {
            var handles=CommentHandles(comment);var tolerance=Math.Max(7,Font.Height/2);
            for(var i=0;i<handles.Length;i++)if(Math.Abs(handles[i].X-point.X)<=tolerance && Math.Abs(handles[i].Y-point.Y)<=tolerance)return i+1;
            return -1;
        }
        private void CancelCommentDrag(){if(proofDragging)normalizedSelection=proofOriginal;proofDragging=false;commentDragging=false;commentPreview=null;picture.Capture=false;picture.Invalidate();}
        private static void DrawHighlight(Graphics graphics, Rectangle bounds, Color color, bool selected)
        {
            if (bounds.Width <= 0 || bounds.Height <= 0) return;
            using (var fill = new SolidBrush(Color.FromArgb(selected ? 52 : 22, color)))
            using (var pen = new Pen(color, selected ? 2.5f : 1f))
            {
                graphics.FillRectangle(fill, bounds);
                graphics.DrawRectangle(pen, bounds);
                if (selected)
                    using (var handle = new SolidBrush(color))
                    {
                        graphics.FillRectangle(handle, bounds.Left - 3, bounds.Top - 3, 6, 6);
                        graphics.FillRectangle(handle, bounds.Right - 3, bounds.Bottom - 3, 6, 6);
                    }
            }
        }

        private void RevealSelection()
        {
            if (!normalizedSelection.HasValue) return;
            var zone = normalizedSelection.Value;
            var x = (int)((zone.X + zone.Width / 2) * picture.Width) + pageBounds[pageIndex].Left;
            var y = (int)((zone.Y + zone.Height / 2) * picture.Height) + pageBounds[pageIndex].Top;
            viewport.AutoScrollPosition = new Point(Math.Max(0, x - viewport.ClientSize.Width / 2), Math.Max(0, y - viewport.ClientSize.Height / 2));
            RefreshVisiblePages();
        }

        private void UpdateNavigationState()
        {
            var pageCount = pdf == null ? (currentPath == null ? 0 : imagePageCount) : pdf.PageCount;
            updatingPageNumber=true;
            try{pageNumber.Maximum=Math.Max(1,pageCount);pageNumber.Value=Math.Max(1,Math.Min(pageCount,pageIndex+1));pageNumber.Enabled=pageCount>0;}
            finally{updatingPageNumber=false;}
            pageLabel.Text=pageCount==0?"Aucun document":" / "+pageCount;
            previousButton.Enabled = pageIndex > 0;
            nextButton.Enabled = pageIndex >= 0 && pageIndex < pageCount - 1;
            zoomOutButton.Enabled = currentImage != null;
            fitButton.Enabled = currentImage != null;
            zoomInButton.Enabled = currentImage != null;
        }

        private static Button CreateToolbarButton(string text, string tooltip, string icon)
        {
            var button = SnipTheme.Button(text, icon);
            button.AccessibleName = tooltip;
            var tip = new ToolTip(); tip.SetToolTip(button, tooltip);
            button.Disposed += (sender, args) => tip.Dispose();
            return button;
        }

        private static Point Clamp(Point point, Rectangle bounds)
        {
            return new Point(
                Math.Max(bounds.Left, Math.Min(bounds.Right, point.X)),
                Math.Max(bounds.Top, Math.Min(bounds.Bottom, point.Y)));
        }

        private static Rectangle NormalizeScreenRectangle(Point first, Point second)
        {
            return Rectangle.FromLTRB(
                Math.Min(first.X, second.X),
                Math.Min(first.Y, second.Y),
                Math.Max(first.X, second.X),
                Math.Max(first.Y, second.Y));
        }

        private void DisposeDocument()
        {
            CancelCommentDrag();selectedCommentId=null;
            dragging = false; panning = false; picture.Capture = false;
            currentPath=null;
            foreach(var surface in pagePictures.Values)
            {
                var image=surface.Image;surface.Image=null;image?.Dispose();
                if(surface!=emptyPicture){pageHost.Controls.Remove(surface);surface.Dispose();}
            }
            pagePictures.Clear();previewSizes.Clear();pageBounds.Clear();picture=emptyPicture;picture.Image=null;picture.Size=Size.Empty;
            pageHost.Size=Size.Empty;
            viewport.AutoScrollMinSize=Size.Empty;viewport.AutoScrollPosition=Point.Empty;
            if (pdf != null) pdf.Dispose();
            pdf = null;
            currentPath = null;
            normalizedSelection = null;
            selectedProofId = null; proofs.Clear(); document = null;
            pageIndex = 0;
            zoomLabel.Text = "100 %";
            pageLabel.Text = "Aucun document";
            UpdateNavigationState();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { DisposeDocument(); popup.Dispose(); zoomMenu.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
