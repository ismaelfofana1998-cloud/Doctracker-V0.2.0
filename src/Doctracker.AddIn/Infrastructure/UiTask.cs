using System;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Doctracker.AddIn.Infrastructure
{
    // Office callbacks do not always install a WinForms SynchronizationContext.
    // Capture the UI dispatcher before awaiting, including fault/cancellation paths.
    internal static class UiTask
    {
        private static SynchronizationContext Capture(Control owner)
        {
            if (owner.IsDisposed || !owner.IsHandleCreated || owner.InvokeRequired)
                throw new InvalidOperationException("Cette opération doit démarrer dans le panneau actif.");
            return new WindowsFormsSynchronizationContext();
        }
        public static Awaitable OnUi(this Task task, Control owner) => new Awaitable(task, Capture(owner));
        public static Awaitable<T> OnUi<T>(this Task<T> task, Control owner) => new Awaitable<T>(task, Capture(owner));

        internal struct Awaitable : INotifyCompletion
        {
            private readonly Task task;
            private readonly SynchronizationContext ui;
            public Awaitable(Task task, SynchronizationContext ui) { this.task=task; this.ui=ui; }
            public Awaitable GetAwaiter() => this;
            public bool IsCompleted => task.IsCompleted;
            public void GetResult() => task.GetAwaiter().GetResult();
            public void OnCompleted(Action continuation)
            {
                var dispatcher=ui;
                task.ConfigureAwait(false).GetAwaiter().OnCompleted(()=>dispatcher.Post(_=>continuation(),null));
            }
        }
        internal struct Awaitable<T> : INotifyCompletion
        {
            private readonly Task<T> task;
            private readonly SynchronizationContext ui;
            public Awaitable(Task<T> task, SynchronizationContext ui) { this.task=task; this.ui=ui; }
            public Awaitable<T> GetAwaiter() => this;
            public bool IsCompleted => task.IsCompleted;
            public T GetResult() => task.GetAwaiter().GetResult();
            public void OnCompleted(Action continuation)
            {
                var dispatcher=ui;
                task.ConfigureAwait(false).GetAwaiter().OnCompleted(()=>dispatcher.Post(_=>continuation(),null));
            }
        }
    }
}
