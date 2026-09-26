using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;

namespace Doctracker.AddIn.Infrastructure
{
    // Small local breadcrumbs, never document text, file names or Excel values.
    internal static class DiagnosticLog
    {
        private static readonly object sync=new object();
        public static string DirectoryPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"Doctracker","Diagnostics");
        private static string LogPath => Path.Combine(DirectoryPath,"session-"+Process.GetCurrentProcess().Id+".log");
        public static void Write(string stage,Exception error=null)
        {
            try
            {
                lock(sync)
                {
                    Directory.CreateDirectory(DirectoryPath);
                    var path=LogPath;
                    if(File.Exists(path) && new FileInfo(path).Length>256*1024)
                    {var previous=path+".previous";if(File.Exists(previous))File.Delete(previous);File.Move(path,previous);}
                    var detail=error==null?"":(" | "+error.GetType().FullName+" 0x"+error.HResult.ToString("X8")+" | "+error.StackTrace);
                    File.AppendAllText(path,DateTime.UtcNow.ToString("O")+" | "+(Environment.Is64BitProcess?"x64":"x86")+" | thread="+Thread.CurrentThread.ManagedThreadId+" "+Thread.CurrentThread.GetApartmentState()+" | "+stage+detail+Environment.NewLine,Encoding.UTF8);
                }
            }
            catch { /* Diagnostics must never break Excel, even on a full or locked disk. */ }
        }
        public static void Start()
        {
            // Retain at most ten previous sessions. No timer and no document backup.
            try
            {
                var files=new DirectoryInfo(DirectoryPath).GetFiles("session-*.log*");
                Array.Sort(files,(a,b)=>b.LastWriteTimeUtc.CompareTo(a.LastWriteTimeUtc));
                for(var i=10;i<files.Length;i++)if(files[i].LastWriteTimeUtc<DateTime.UtcNow.AddDays(-1))try{files[i].Delete();}catch{}
            }
            catch{}
            Write("Start "+typeof(DiagnosticLog).Assembly.GetName().Version);
        }
    }
}
