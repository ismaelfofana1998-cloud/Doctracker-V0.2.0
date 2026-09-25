using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Doctracker.Core.Models;
using PdfiumViewer;

namespace Doctracker.AddIn.Infrastructure
{
    public sealed class OcrWorkRequest
    {
        public string Source {get;set;}
        public int PageNumber {get;set;}=1;
        public float X {get;set;}
        public float Y {get;set;}
        public float Width {get;set;}=1;
        public float Height {get;set;}=1;
        public bool Table {get;set;}
    }
    internal static class OcrWorkerEntry
    {
        public static int Run(string requestPath,string responsePath)
        {
            try
            {
                OcrWorkRequest request;
                using(var reader=XmlReader.Create(requestPath,ReaderSettings()))request=(OcrWorkRequest)new XmlSerializer(typeof(OcrWorkRequest)).Deserialize(reader);
                if(request==null || request.PageNumber<1 || !Valid(request.X) || !Valid(request.Y) || !Valid(request.Width) || !Valid(request.Height) ||
                    request.Width<=0 || request.Height<=0 || request.X+request.Width>1.00001 || request.Y+request.Height>1.00001)throw new InvalidDataException();
                Console.Out.WriteLine("WorkerCropStart");
                using(var crop=ReadCrop(request))
                using(var engine=new TesseractOcrEngine())
                {
                    Console.Out.WriteLine("WorkerCropReady");
                    Console.Out.WriteLine("WorkerOcrStart");
                    var page=engine.Recognize(crop,request.Table);page.PageNumber=request.PageNumber;
                    Console.Out.WriteLine("WorkerOcrReady");
                    page.Text=Clean(page.Text);foreach(var word in page.Words)word.Text=Clean(word.Text);
                    using(var writer=XmlWriter.Create(responsePath,new XmlWriterSettings {Encoding=new UTF8Encoding(false)}))
                        new XmlSerializer(typeof(PageTextRecord)).Serialize(writer,page);
                    Console.Out.WriteLine("WorkerResultReady");
                    return 0;
                }
            }
            catch {Console.Out.WriteLine("WorkerFailure");return 70;}
        }
        internal static XmlReaderSettings ReaderSettings()=>new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=16*1024*1024};
        private static bool Valid(float value)=>!float.IsNaN(value) && !float.IsInfinity(value) && value>=0 && value<=1;
        private static Bitmap ReadCrop(OcrWorkRequest request)
        {
            if(Path.GetExtension(request.Source).Equals(".pdf",StringComparison.OrdinalIgnoreCase))
            {
                NativePdfiumLoader.EnsureLoaded();
                using(var pdf=PdfDocument.Load(request.Source))
                {
                    if(request.PageNumber>pdf.PageCount)throw new InvalidDataException();
                    var size=DocumentIndexer.RenderSize(pdf.PageSizes[request.PageNumber-1]);
                    using(var source=pdf.Render(request.PageNumber-1,size.Width,size.Height,144,144,PdfRenderFlags.Annotations))return Crop(source,request);
                }
            }
            using(var source=Image.FromFile(request.Source))
            {
                var count=DocumentIndexer.ImagePageCount(source);if(request.PageNumber>count)throw new InvalidDataException();
                if(count>1)source.SelectActiveFrame(FrameDimension.Page,request.PageNumber-1);
                return Crop(source,request);
            }
        }
        private static Bitmap Crop(Image source,OcrWorkRequest request)
        {
            var box=Rectangle.FromLTRB(Math.Max(0,(int)Math.Floor(request.X*source.Width)),Math.Max(0,(int)Math.Floor(request.Y*source.Height)),
                Math.Min(source.Width,(int)Math.Ceiling((request.X+request.Width)*source.Width)),Math.Min(source.Height,(int)Math.Ceiling((request.Y+request.Height)*source.Height)));
            if(box.Width<2 || box.Height<2)throw new InvalidDataException();
            var scale=Math.Min(1,3000d/Math.Max(box.Width,box.Height));
            var crop=new Bitmap(Math.Max(1,(int)Math.Ceiling(box.Width*scale)),Math.Max(1,(int)Math.Ceiling(box.Height*scale)));
            try {using(var graphics=Graphics.FromImage(crop))graphics.DrawImage(source,new Rectangle(Point.Empty,crop.Size),box,GraphicsUnit.Pixel);return crop;}
            catch {crop.Dispose();throw;}
        }
        private static string Clean(string value)
        {
            if(string.IsNullOrEmpty(value))return value;
            var result=new StringBuilder(value.Length);
            for(var i=0;i<value.Length;i++)
            {
                var c=value[i];
                if(char.IsHighSurrogate(c) && i+1<value.Length && char.IsLowSurrogate(value[i+1])){result.Append(c).Append(value[++i]);continue;}
                if(XmlConvert.IsXmlChar(c))result.Append(c);
            }
            return result.ToString();
        }
    }
}
