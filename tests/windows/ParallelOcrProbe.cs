using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Doctracker.Core.Models;
using Doctracker.Core.Services;

public static class ParallelOcrProbe
{
    private const BindingFlags Flags=BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic;
    private static void Check(bool value,string message){if(!value)throw new Exception(message);}
    private static object Call(object target,string method,params object[] args)
    {
        try{return target.GetType().GetMethod(method,Flags).Invoke(target,args);}
        catch(TargetInvocationException failure){throw failure.InnerException;}
    }
    private static void Mode(string worker,string mode)
    {
        var folder=Path.Combine(Path.GetDirectoryName(worker),"events");
        if(Directory.Exists(folder))Directory.Delete(folder,true);Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(worker),"mode.txt"),mode);
    }
    private static string[] Starts(string worker){return Directory.GetFiles(Path.Combine(Path.GetDirectoryName(worker),"events"),"*.start");}
    private static void CheckPool(string worker,bool cancelled)
    {
        var events=new List<KeyValuePair<long,int>>();
        foreach(var start in Starts(worker))
        {
            var pid=int.Parse(Path.GetFileNameWithoutExtension(start));
            try{using(var process=Process.GetProcessById(pid))Check(process.HasExited,"OCR child survived: "+pid);}
            catch(ArgumentException){}
            events.Add(new KeyValuePair<long,int>(long.Parse(File.ReadAllText(start)),1));
            var end=Path.ChangeExtension(start,"end");
            if(File.Exists(end))events.Add(new KeyValuePair<long,int>(long.Parse(File.ReadAllText(end)),-1));
            else Check(cancelled,"Worker completion marker missing");
        }
        if(!cancelled)
        {
            var active=0;var maximum=0;
            foreach(var item in events.OrderBy(e=>e.Key)){active+=item.Value;maximum=Math.Max(maximum,active);}
            Check(maximum==2,"Expected exactly two overlapping OCR workers, observed "+maximum);
            Check(active==0,"Pool did not drain");
        }
    }
    public static string Run(Assembly assembly,string temp,string pdf,string worker)
    {
        var ocrType=assembly.GetType("Doctracker.AddIn.Infrastructure.IsolatedOcrEngine",true);
        var indexerType=assembly.GetType("Doctracker.AddIn.Infrastructure.DocumentIndexer",true);
        var engine=ocrType.GetConstructor(Flags,null,new[]{typeof(string),typeof(int)},null).Invoke(new object[]{worker,10000});
        var store=new ProjectStore(Path.Combine(temp,"pool-project"));
        var state=new ProjectState();var importer=new DocumentImporter(store);
        var indexer=indexerType.GetConstructors(Flags)[0].Invoke(new[]{(object)store,engine});
        indexerType.GetProperty("WorkerCount").SetValue(indexer,2,null);
        try
        {
            for(var i=0;i<7;i++)
            {
                var path=Path.Combine(temp,"pool-"+i+".png");
                using(var bitmap=new Bitmap(40+i,40))bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
                importer.Import(state,path,"pool");
            }
            var scope=state.Documents.Take(6).ToList();
            Mode(worker,"normal");
            var errors=(List<string>)Call(indexer,"IndexMissing",state,null,CancellationToken.None,scope,true,false,true);
            Check(errors.Count==0 && scope.All(d=>d.IndexComplete) && !state.Documents[6].IndexComplete,"Pool lost documents or ignored selected scope");
            Check(Starts(worker).Length==6,"Unexpected number of worker launches");CheckPool(worker,false);
            Check(store.LoadOrCreate("").Documents.Count(d=>d.IndexComplete)==6,"Concurrent indexing metadata not durable");

            Mode(worker,"fail-source");
            var failureState=new ProjectState();var failureStore=new ProjectStore(Path.Combine(temp,"pool-failure"));
            var failureImporter=new DocumentImporter(failureStore);
            for(var i=0;i<4;i++)failureImporter.Import(failureState,Path.Combine(temp,"pool-"+i+".png"),"pool");
            File.WriteAllText(Path.Combine(Path.GetDirectoryName(worker),"fail-source.txt"),failureStore.ResolveDocumentPath(failureState.Documents[0]));
            var failureIndexer=indexerType.GetConstructors(Flags)[0].Invoke(new[]{(object)failureStore,engine});
            indexerType.GetProperty("WorkerCount").SetValue(failureIndexer,2,null);
            errors=(List<string>)Call(failureIndexer,"IndexMissing",failureState,null,CancellationToken.None,null,true,false,true);
            Check(errors.Count==1 && !failureState.Documents[0].IndexComplete && failureState.Documents.Skip(1).All(d=>d.IndexComplete),"One failed document cancelled unrelated documents");
            CheckPool(worker,false);

            var pdfDoc=importer.Import(state,pdf,"pool");
            Mode(worker,"normal");
            Call(indexer,"Index",state,pdfDoc,null,CancellationToken.None,true,true);
            Check(pdfDoc.IndexComplete && pdfDoc.PageCount==30 && pdfDoc.IndexedPages.Select(p=>p.PageNumber).SequenceEqual(Enumerable.Range(1,30)),"Parallel page results lost their source order");
            Check(pdfDoc.IndexedPages.All(p=>p.Text=="PAGE "+p.PageNumber),"Parallel page identities mixed");CheckPool(worker,false);
            var key=pdfDoc.IndexKey;
            Mode(worker,"fail-page");var failed=false;
            try{Call(indexer,"Index",state,pdfDoc,null,CancellationToken.None,true,true);}catch(InvalidOperationException){failed=true;}
            Check(failed && pdfDoc.IndexComplete && pdfDoc.IndexKey==key,"Failed parallel reindex replaced the previous index");
            CheckPool(worker,true);

            Mode(worker,"hang");
            using(var cancel=new CancellationTokenSource())
            {
                var task=Task.Run(()=>{try{Call(indexer,"Index",state,pdfDoc,null,cancel.Token,true,true);return false;}catch(OperationCanceledException){return true;}});
                var deadline=Stopwatch.StartNew();
                while(Starts(worker).Length<2 && deadline.ElapsedMilliseconds<10000)Thread.Sleep(20);
                cancel.Cancel();Check(task.Wait(5000) && task.Result,"Pool cancellation did not finish promptly");
                Check(Starts(worker).Length==2,"Cancellation probe did not start two children");
            }
            Check(pdfDoc.IndexKey==key && pdfDoc.IndexComplete,"Cancelled parallel reindex replaced the previous index");CheckPool(worker,true);
            return "PASS: bounded overlapping OCR workers, selected documents, page order, durable saves, isolated failure and pool cancellation";
        }
        finally{((IDisposable)engine).Dispose();}
    }
}
