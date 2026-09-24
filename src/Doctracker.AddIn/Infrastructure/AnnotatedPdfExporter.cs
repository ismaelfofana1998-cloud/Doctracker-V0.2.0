using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using Doctracker.AddIn.UI;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using PdfSharp.Drawing;

namespace Doctracker.AddIn.Infrastructure
{
    internal static class AnnotatedPdfExporter
    {
        // Export is a flattened visual copy; original source bytes are never changed.
        public static void Export(string source, DocumentRecord document, SnipRecord[] snips, string destination, CancellationToken cancellation)
        {
            var temp=destination+"."+Guid.NewGuid().ToString("N")+".tmp";
            try
            {
                using(var output=new PdfSharp.Pdf.PdfDocument())
                {
                    if(Path.GetExtension(source).Equals(".pdf",StringComparison.OrdinalIgnoreCase))
                    {
                        NativePdfiumLoader.EnsureLoaded();
                        using(var pdf=PdfiumViewer.PdfDocument.Load(source))
                        for(var i=0;i<pdf.PageCount;i++)
                        {
                            cancellation.ThrowIfCancellationRequested();var size=DocumentIndexer.RenderSize(pdf.PageSizes[i]);
                            using(var rendered=pdf.Render(i,size.Width,size.Height,144,144,PdfiumViewer.PdfRenderFlags.Annotations))
                            using(var bitmap=new Bitmap(rendered))AddPage(output,bitmap,pdf.PageSizes[i],document,snips,i+1);
                        }
                    }
                    else
                    {
                        using(var original=Image.FromFile(source))
                        for(var i=0;i<DocumentIndexer.ImagePageCount(original);i++)
                        {
                            cancellation.ThrowIfCancellationRequested();if(DocumentIndexer.ImagePageCount(original)>1)original.SelectActiveFrame(FrameDimension.Page,i);
                            using(var bitmap=new Bitmap(original))AddPage(output,bitmap,new SizeF(bitmap.Width*72f/144,bitmap.Height*72f/144),document,snips,i+1);
                        }
                    }
                    output.Info.Title=document.DisplayName;output.Info.Subject="Doctracker — copie annotée, original conservé";
                    output.Save(temp);
                }
                cancellation.ThrowIfCancellationRequested();
                if(File.Exists(destination))File.Replace(temp,destination,destination+".bak");else File.Move(temp,destination);
            }
            finally{if(File.Exists(temp))File.Delete(temp);}
        }
        private static void AddPage(PdfSharp.Pdf.PdfDocument output,Bitmap bitmap,SizeF size,DocumentRecord document,SnipRecord[] snips,int number)
        {
            using(var graphics=Graphics.FromImage(bitmap))
            using(var font=new Font("Segoe UI",Math.Max(9,bitmap.Width/110f),FontStyle.Bold,GraphicsUnit.Pixel))
            {
                DocumentOverlay.Draw(graphics, bitmap.Size, document, number);
                foreach(var snip in snips.Where(s=>s.PageNumber==number))
                {
                    var color=SnipTheme.ColorFor(snip.SourceType??snip.Type);
                    var rect=new RectangleF((float)snip.X*bitmap.Width,(float)snip.Y*bitmap.Height,(float)snip.Width*bitmap.Width,(float)snip.Height*bitmap.Height);
                    using(var fill=new SolidBrush(Color.FromArgb(35,color)))using(var pen=new Pen(color,Math.Max(2,bitmap.Width/600f)))
                    {graphics.FillRectangle(fill,rect);graphics.DrawRectangle(pen,rect.X,rect.Y,rect.Width,rect.Height);}
                    var label=CrossReferences.SnipLabel(document,snip);
                    var labelSize=graphics.MeasureString(label,font);var x=Math.Max(0,Math.Min(rect.Left,bitmap.Width-labelSize.Width));var y=Math.Max(0,rect.Top-labelSize.Height);
                    graphics.FillRectangle(Brushes.White,x,y,labelSize.Width,labelSize.Height);using(var brush=new SolidBrush(color))graphics.DrawString(label,font,brush,x,y);
                }
            }
            var page=output.AddPage();page.Width=XUnit.FromPoint(size.Width);page.Height=XUnit.FromPoint(size.Height);
            using(var image=XImage.FromGdiPlusImage(bitmap))using(var graphics=XGraphics.FromPdfPage(page))graphics.DrawImage(image,0,0,size.Width,size.Height);
        }
    }
}
