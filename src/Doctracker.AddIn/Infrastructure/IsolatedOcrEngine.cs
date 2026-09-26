using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Serialization;
using Doctracker.Core.Models;

namespace Doctracker.AddIn.Infrastructure
{
    internal sealed class IsolatedOcrEngine : IOcrEngine
    {
        private readonly object sync=new object();
        private readonly SemaphoreSlim serial=new SemaphoreSlim(1,1);
        private readonly string workerOverride;
        private readonly int timeoutMilliseconds;
        private Process active;
        private bool disposed;
        private bool batchEngine;
        internal IsolatedOcrEngine CreateBatchEngine()=>new IsolatedOcrEngine(workerOverride,timeoutMilliseconds) {batchEngine=true};
        public CancellationToken Cancellation {get;set;}
        public IsolatedOcrEngine():this(null,90000){}
        // Dependency injection also permits testing child failure without crashing Excel.
        internal IsolatedOcrEngine(string workerExecutable,int timeoutMilliseconds)
        {workerOverride=workerExecutable;this.timeoutMilliseconds=timeoutMilliseconds;}
        public PageTextRecord Recognize(Bitmap bitmap,bool table=false)
        {
            if(bitmap==null)throw new ArgumentNullException(nameof(bitmap));
            var folder=NewWorkDirectory();
            try
            {
                Cancellation.ThrowIfCancellationRequested();
                var path=Path.Combine(folder,"pixels.png");bitmap.Save(path,ImageFormat.Png);
                return Execute<PageTextRecord>(new OcrWorkRequest {Source=path,Table=table},folder,Cancellation);
            }
            finally {RemoveWorkDirectory(folder);}
        }
        public PageTextRecord RecognizeRegion(string path,int page,RectangleF rectangle,bool table,CancellationToken cancellation)
        {
            var folder=NewWorkDirectory();
            try {return Execute<PageTextRecord>(new OcrWorkRequest {Source=path,PageNumber=page,X=rectangle.X,Y=rectangle.Y,Width=rectangle.Width,Height=rectangle.Height,Table=table},folder,cancellation);}
            finally {RemoveWorkDirectory(folder);}
        }
        public List<PageTextRecord> RecognizePages(string path,IEnumerable<int> pages,Action<int> progress,CancellationToken cancellation)
        {
            var requested=pages.ToList();
            if(requested.Count==0)return new List<PageTextRecord>();
            if(requested.Count>16 || requested.Any(n=>n<1) || requested.Distinct().Count()!=requested.Count)
                throw new ArgumentException("Lot OCR invalide (1 à 16 pages distinctes).");
            var folder=NewWorkDirectory();
            try
            {
                var result=Execute<OcrBatchResult>(new OcrWorkRequest {Source=path,Pages=requested},folder,cancellation,progress);
                if(result.Pages==null || !result.Pages.Select(p=>p.PageNumber).SequenceEqual(requested))
                    throw new InvalidDataException("Le lot OCR reçu est incomplet.");
                return result.Pages;
            }
            finally {RemoveWorkDirectory(folder);}
        }
        private T Execute<T>(OcrWorkRequest request,string folder,CancellationToken cancellation,Action<int> progress=null) where T:class
        {
            serial.Wait(cancellation);
            try
            {
                var executable=workerOverride??NativePdfiumLoader.GetDeploymentDirectories().Select(root=>Path.Combine(root,"Doctracker.OcrWorker."+(Environment.Is64BitProcess?"x64":"x86")+".exe")).FirstOrDefault(File.Exists);
                if(string.IsNullOrEmpty(executable) || !File.Exists(executable))throw new FileNotFoundException("Le moteur OCR isolé est absent. Réinstallez le ZIP complet de Doctracker.");
                var input=Path.Combine(folder,"request.xml");var output=Path.Combine(folder,"response.xml");
                using(var writer=XmlWriter.Create(input))new XmlSerializer(typeof(OcrWorkRequest)).Serialize(writer,request);
                long lastProgress=Stopwatch.GetTimestamp();
                using(var process=new Process {StartInfo=new ProcessStartInfo(executable) {
                    Arguments=Quote(input)+" "+Quote(output)+" "+Process.GetCurrentProcess().Id,
                    WorkingDirectory=Path.GetDirectoryName(executable),UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true}})
                {
                    if(batchEngine)process.StartInfo.EnvironmentVariables["OMP_THREAD_LIMIT"]="1";
                    process.OutputDataReceived+=(sender,args)=>{
                        if(args.Data!=null && args.Data.StartsWith("WorkerPageReady ") && int.TryParse(args.Data.Substring(16),out var pageNumber) && request.Pages.Contains(pageNumber))
                        {
                            Interlocked.Exchange(ref lastProgress,Stopwatch.GetTimestamp());
                            DiagnosticLog.Write("WorkerPageReady page="+pageNumber);
                            try {progress?.Invoke(pageNumber);}catch { /* Progress must not crash the output reader. */ }
                        }
                        if(args.Data!=null && (args.Data.StartsWith("WorkerRenderMs ") || args.Data.StartsWith("WorkerRecognitionMs ")) &&
                            long.TryParse(args.Data.Substring(args.Data.IndexOf(' ')+1),out var milliseconds))DiagnosticLog.Write(args.Data);
                        if(new[]{"WorkerCropStart","WorkerCropReady","WorkerOcrStart","WorkerOcrReady","WorkerResultReady","WorkerFailure"}.Contains(args.Data))DiagnosticLog.Write(args.Data);
                    };
                    process.ErrorDataReceived+=(sender,args)=>{}; // Drain native warnings; never log recognized content.
                    try
                    {
                        cancellation.ThrowIfCancellationRequested();
                        lock(sync)
                        {
                            if(disposed)throw new ObjectDisposedException(nameof(IsolatedOcrEngine));
                            if(!process.Start())throw new InvalidOperationException("Le processus OCR n'a pas pu démarrer.");
                            active=process;
                        }
                        DiagnosticLog.Write("WorkerStarted pid="+process.Id);
                        process.BeginOutputReadLine();process.BeginErrorReadLine();
                        Interlocked.Exchange(ref lastProgress,Stopwatch.GetTimestamp());
                        while(!process.WaitForExit(100))
                        {
                            cancellation.ThrowIfCancellationRequested();
                            if((Stopwatch.GetTimestamp()-Interlocked.Read(ref lastProgress))*1000.0/Stopwatch.Frequency>timeoutMilliseconds)throw new TimeoutException("La reconnaissance a dépassé le délai autorisé. Réduisez la zone ou essayez une autre page.");
                        }
                        process.WaitForExit(); // Complete stdout callbacks before returning the result.
                        cancellation.ThrowIfCancellationRequested();
                        if(process.ExitCode!=0 || !File.Exists(output))
                        {
                            DiagnosticLog.Write("WorkerExit code=0x"+process.ExitCode.ToString("X8"));
                            throw new InvalidOperationException("Le moteur de reconnaissance s'est arrêté. Ce traitement n'a pas été appliqué. Relancez sur une sélection plus petite ou transmettez le journal de diagnostic.");
                        }
                        using(var reader=XmlReader.Create(output,OcrWorkerEntry.ReaderSettings()))
                        {
                            var page=(T)new XmlSerializer(typeof(T)).Deserialize(reader);
                            if(page==null)throw new InvalidDataException("Réponse OCR vide.");
                            return page;
                        }
                    }
                    finally
                    {
                        lock(sync){if(ReferenceEquals(active,process))active=null;}
                        Stop(process);
                    }
                }
            }
            finally {serial.Release();}
        }
        private static string Quote(string path)=>"\""+path+"\""; // Generated file paths never end in a backslash or contain quotes.
        private static string NewWorkDirectory()
        {
            var path=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"Doctracker","OcrWork",Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);return path;
        }
        private static void RemoveWorkDirectory(string path)
        {try {Directory.Delete(path,true);}catch(IOException){}catch(UnauthorizedAccessException){}}
        private static void Stop(Process process)
        {try {if(!process.HasExited){process.Kill();process.WaitForExit(2000);}}catch(InvalidOperationException){}catch(System.ComponentModel.Win32Exception){}}
        public void Dispose()
        {
            lock(sync)
            {
                disposed=true;
                if(active!=null)try{if(!active.HasExited)active.Kill();}catch(InvalidOperationException){}catch(System.ComponentModel.Win32Exception){}
            }
        }
    }
}
