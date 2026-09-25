using System;
using System.Diagnostics;
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
                return Execute(new OcrWorkRequest {Source=path,Table=table},folder,Cancellation);
            }
            finally {RemoveWorkDirectory(folder);}
        }
        public PageTextRecord RecognizeRegion(string path,int page,RectangleF rectangle,bool table,CancellationToken cancellation)
        {
            var folder=NewWorkDirectory();
            try {return Execute(new OcrWorkRequest {Source=path,PageNumber=page,X=rectangle.X,Y=rectangle.Y,Width=rectangle.Width,Height=rectangle.Height,Table=table},folder,cancellation);}
            finally {RemoveWorkDirectory(folder);}
        }
        private PageTextRecord Execute(OcrWorkRequest request,string folder,CancellationToken cancellation)
        {
            serial.Wait(cancellation);
            try
            {
                var executable=workerOverride??NativePdfiumLoader.GetDeploymentDirectories().Select(root=>Path.Combine(root,"Doctracker.OcrWorker."+(Environment.Is64BitProcess?"x64":"x86")+".exe")).FirstOrDefault(File.Exists);
                if(string.IsNullOrEmpty(executable) || !File.Exists(executable))throw new FileNotFoundException("Le moteur OCR isolé est absent. Réinstallez le ZIP complet de Doctracker.");
                var input=Path.Combine(folder,"request.xml");var output=Path.Combine(folder,"response.xml");
                using(var writer=XmlWriter.Create(input))new XmlSerializer(typeof(OcrWorkRequest)).Serialize(writer,request);
                using(var process=new Process {StartInfo=new ProcessStartInfo(executable) {
                    Arguments=Quote(input)+" "+Quote(output)+" "+Process.GetCurrentProcess().Id,
                    WorkingDirectory=Path.GetDirectoryName(executable),UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true}})
                {
                    process.OutputDataReceived+=(sender,args)=>{
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
                        var elapsed=Stopwatch.StartNew();
                        while(!process.WaitForExit(100))
                        {
                            cancellation.ThrowIfCancellationRequested();
                            if(elapsed.ElapsedMilliseconds>timeoutMilliseconds)throw new TimeoutException("La reconnaissance a dépassé le délai autorisé. Réduisez la zone ou essayez une autre page.");
                        }
                        process.WaitForExit(); // Complete stdout callbacks before returning the result.
                        cancellation.ThrowIfCancellationRequested();
                        if(process.ExitCode!=0 || !File.Exists(output))
                        {
                            DiagnosticLog.Write("WorkerExit code=0x"+process.ExitCode.ToString("X8"));
                            throw new InvalidOperationException("Le moteur de reconnaissance s'est arrêté. Excel est resté ouvert et ce snip n'a pas été inséré. Essayez une zone plus petite ou transmettez le journal de diagnostic.");
                        }
                        using(var reader=XmlReader.Create(output,OcrWorkerEntry.ReaderSettings()))
                        {
                            var page=(PageTextRecord)new XmlSerializer(typeof(PageTextRecord)).Deserialize(reader);
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
