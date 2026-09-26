using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

// Deterministic worker for testing the real process pool, not an OCR replacement.
public static class ParallelFakeWorker
{
    public static int Main(string[] args)
    {
        var root=AppDomain.CurrentDomain.BaseDirectory;
        var events=Path.Combine(root,"events");Directory.CreateDirectory(events);
        var marker=Path.Combine(events,Process.GetCurrentProcess().Id.ToString());
        File.WriteAllText(marker+".start",DateTime.UtcNow.Ticks.ToString());
        try
        {
            if(Environment.GetEnvironmentVariable("OMP_THREAD_LIMIT")!="1")return 91;
            var request=new XmlDocument();request.Load(args[0]);
            var source=request.SelectSingleNode("/OcrWorkRequest/Source").InnerText;
            var pages=request.SelectNodes("/OcrWorkRequest/Pages/int");
            var mode=File.ReadAllText(Path.Combine(root,"mode.txt"));
            if(mode=="hang"){Thread.Sleep(30000);return 92;}
            // The first two workers must actually overlap, not merely be started in sequence.
            var deadline=Stopwatch.StartNew();
            while(Directory.GetFiles(events,"*.start").Length<2 && deadline.ElapsedMilliseconds<5000)Thread.Sleep(10);
            if(Directory.GetFiles(events,"*.start").Length<2)return 93;
            Thread.Sleep(int.Parse(pages[0].InnerText)%2==0?150:350);
            if(mode=="fail-source" && source==File.ReadAllText(Path.Combine(root,"fail-source.txt")))return 71;
            var response=new StringBuilder("<OcrBatchResult><Pages>");
            foreach(XmlNode page in pages)
            {
                if(mode=="fail-page" && page.InnerText=="17")return 72;
                response.Append("<PageTextRecord PageNumber='").Append(page.InnerText).Append("'>PAGE ").Append(page.InnerText).Append("</PageTextRecord>");
                Console.WriteLine("WorkerPageReady "+page.InnerText);
            }
            response.Append("</Pages></OcrBatchResult>");File.WriteAllText(args[1],response.ToString());return 0;
        }
        finally {File.WriteAllText(marker+".end",DateTime.UtcNow.Ticks.ToString());}
    }
}
