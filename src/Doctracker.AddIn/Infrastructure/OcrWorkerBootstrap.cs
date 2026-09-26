// Built separately as x86/x64 executables; never compiled into the Excel add-in.
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

internal static class OcrWorkerBootstrap
{
    [DllImport("kernel32.dll")] private static extern uint SetErrorMode(uint mode);
    [DllImport("kernel32.dll")] private static extern IntPtr OpenProcess(uint access,bool inherit,int pid);
    [DllImport("kernel32.dll")] private static extern uint WaitForSingleObject(IntPtr handle,uint milliseconds);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
    private static int Main(string[] args)
    {
        SetErrorMode(0x0001|0x0002); // A child crash must not leave a modal Windows error dialog.
        if(args.Length!=3)return 64;
        int parentId;if(!int.TryParse(args[2],out parentId))return 64;
        var parent=OpenProcess(0x00100000,false,parentId);
        try
        {
            // The host bounds each page by an inactivity deadline and kills on cancellation.
            using(var watch=new Timer(_=>{if(parent!=IntPtr.Zero && WaitForSingleObject(parent,0)==0)Environment.Exit(125);},null,2000,2000))
            {
                var assembly=Assembly.LoadFrom(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Doctracker.AddIn.dll"));
                var entry=assembly.GetType("Doctracker.AddIn.Infrastructure.OcrWorkerEntry",true).GetMethod("Run",BindingFlags.Public|BindingFlags.Static);
                return (int)entry.Invoke(null,new object[]{args[0],args[1]});
            }
        }
        catch {Console.Out.WriteLine("WorkerFailure");return 70;}
        finally {if(parent!=IntPtr.Zero)CloseHandle(parent);}
    }
}
