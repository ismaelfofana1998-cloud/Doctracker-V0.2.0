using System;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

// Calls the awaiter in the built add-in, with the ambient context deliberately absent/wrong.
public static class UiContinuationProbe
{
    public static Task<string> Run(Assembly assembly, Control owner, string outcome, bool generic)
    {
        int uiThread=Thread.CurrentThread.ManagedThreadId;
        var original=SynchronizationContext.Current;
        var source=new TaskCompletionSource<int>();
        var result=new TaskCompletionSource<string>();
        try
        {
            SynchronizationContext.SetSynchronizationContext(generic ? null : new SynchronizationContext());
            var method=assembly.GetType("Doctracker.AddIn.Infrastructure.UiTask",true)
                .GetMethods(BindingFlags.Public|BindingFlags.Static)
                .Single(m=>m.Name=="OnUi" && m.IsGenericMethodDefinition==generic);
            if(generic)method=method.MakeGenericMethod(typeof(int));
            if(outcome=="completed")source.SetResult(42);
            object awaiter=method.Invoke(null,new object[]{source.Task,owner});
            var getResult=awaiter.GetType().GetMethod("GetResult");
            Action resume=()=>{
                try
                {
                    if(Thread.CurrentThread.ManagedThreadId!=uiThread || owner.InvokeRequired)
                        throw new Exception("Continuation escaped the UI thread: "+outcome);
                    try
                    {
                        object value=getResult.Invoke(awaiter,null);
                        if(outcome=="fault" || outcome=="cancel")throw new Exception("Task failure lost.");
                        if(generic && !Equals(value,42))throw new Exception("Task result lost.");
                    }
                    catch(TargetInvocationException e)
                    {
                        if(outcome=="fault" && e.InnerException is FormatException) { }
                        else if(outcome=="cancel" && e.InnerException is OperationCanceledException) { }
                        else throw;
                    }
                    result.SetResult("PASS: UI continuation "+outcome+" generic="+generic);
                }
                catch(Exception e){result.SetException(e);}
            };
            bool completed=(bool)awaiter.GetType().GetProperty("IsCompleted").GetValue(awaiter,null);
            if(completed)resume();
            else
            {
                ((INotifyCompletion)awaiter).OnCompleted(resume);
                ThreadPool.QueueUserWorkItem(_=>{
                    if(outcome=="fault")source.SetException(new FormatException("Synthetic extraction error"));
                    else if(outcome=="cancel")source.SetCanceled();
                    else source.SetResult(42);
                });
            }
        }
        finally {SynchronizationContext.SetSynchronizationContext(original);}
        return result.Task;
    }
}
