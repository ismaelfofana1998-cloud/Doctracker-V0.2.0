using System;
using System.IO;
using System.Runtime.InteropServices;
using Doctracker.Core.Services;

namespace Doctracker.AddIn.Infrastructure
{
    internal static class OcrSettings
    {
        private static readonly object sync=new object();
        public static string SettingsPath=>Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"Doctracker","ocr-workers.txt");
        public static int Maximum=>OcrConcurrencyPolicy.Maximum(Environment.ProcessorCount);
        public static int Recommended=>OcrConcurrencyPolicy.Recommended(Environment.ProcessorCount,MemoryGb);
        public static double MemoryGb
        {
            get {var memory=new MemoryStatus {Length=(uint)Marshal.SizeOf(typeof(MemoryStatus))};return GlobalMemoryStatusEx(ref memory)?memory.TotalPhysical/1073741824d:0;}
        }
        public static int LoadWorkers()
        {
            lock(sync)
            {
                try {if(int.TryParse(File.ReadAllText(SettingsPath).Trim(),out var count) && count>0)return OcrConcurrencyPolicy.Clamp(count,Environment.ProcessorCount);}
                catch(IOException){}catch(UnauthorizedAccessException){}
                return Recommended;
            }
        }
        public static void SaveWorkers(int count)
        {
            lock(sync)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath));
                var temporary=SettingsPath+"."+Guid.NewGuid().ToString("N")+".tmp";
                try
                {
                    File.WriteAllText(temporary,OcrConcurrencyPolicy.Clamp(count,Environment.ProcessorCount).ToString());
                    if(File.Exists(SettingsPath))File.Replace(temporary,SettingsPath,null);else File.Move(temporary,SettingsPath);
                }
                finally {if(File.Exists(temporary))File.Delete(temporary);}
            }
        }
        [DllImport("kernel32.dll",SetLastError=true)][return:MarshalAs(UnmanagedType.Bool)]
        private static extern bool GlobalMemoryStatusEx(ref MemoryStatus status);
        [StructLayout(LayoutKind.Sequential)] private struct MemoryStatus
        {
            public uint Length,Load;
            public ulong TotalPhysical,AvailablePhysical,TotalPageFile,AvailablePageFile,TotalVirtual,AvailableVirtual,AvailableExtendedVirtual;
        }
    }
}
