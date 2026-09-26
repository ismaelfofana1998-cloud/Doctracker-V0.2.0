using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Doctracker.AddIn.Infrastructure
{
    internal static class WorksheetKeyboardFocus
    {
        // Select/Activate through COM does not guarantee native keyboard focus in a
        // VSTO task pane. Hand it to the worksheet synchronously, without SendKeys.
        // SetFocus: https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-setfocus
        public static bool ReturnToWorksheet(Control pane,IntPtr window,bool allowEmptyFocus=false)
            =>ReturnToChild(pane,window,"EXCEL7",allowEmptyFocus);

        // A queued selection refresh must not run over F2, formula-bar editing,
        // ribbon navigation or another app. Remember the exact worksheet HWND
        // (including split panes), not merely Excel's selected COM range.
        public static IntPtr CaptureWorksheet(IntPtr window)=>CaptureChild(window,"EXCEL7");
        internal static IntPtr CaptureChild(IntPtr window,string childClass)
        {
            var focus=GetFocus();
            return window!=IntPtr.Zero && focus!=IntPtr.Zero &&
                GetForegroundWindow()==GetAncestor(window,2) && IsInside(window,focus) &&
                IsWindowVisible(focus) && IsWindowEnabled(focus) && ClassName(focus)==childClass
                ? focus : IntPtr.Zero;
        }

        public static bool RestoreAfterPaneRefresh(Control pane,IntPtr window,IntPtr originalFocus)
        {
            if(pane==null || pane.IsDisposed || !pane.IsHandleCreated || pane.InvokeRequired ||
                window==IntPtr.Zero || originalFocus==IntPtr.Zero || !IsInside(window,originalFocus) ||
                IsInside(pane.Handle,originalFocus) || !IsWindowVisible(originalFocus) || !IsWindowEnabled(originalFocus) ||
                GetWindowThreadProcessId(originalFocus,IntPtr.Zero)!=GetCurrentThreadId())return false;
            var root=GetAncestor(window,2);
            if(GetForegroundWindow()!=root || !IsWindowEnabled(root))return false;
            var focus=GetFocus();
            if(focus==originalFocus)return true;
            // Only undo focus lost to this synchronous pane refresh. Never force
            // the grid over an editor/dialog or queue a later focus callback.
            if(focus!=IntPtr.Zero && !IsInside(pane.Handle,focus))return false;
            SetFocus(originalFocus);
            return GetFocus()==originalFocus;
        }

        // Class-name injection permits exercising the actual Win32 handoff without Office.
        internal static bool ReturnToChild(Control pane,IntPtr window,string childClass,bool allowEmptyFocus)
        {
            if(pane==null || pane.IsDisposed || !pane.IsHandleCreated || pane.InvokeRequired || window==IntPtr.Zero)return false;
            var root=GetAncestor(window,2);
            if(root==IntPtr.Zero || GetForegroundWindow()!=root || !IsWindowEnabled(root))return false;
            var focus=GetFocus();
            if(focus!=IntPtr.Zero && IsInside(window,focus) && !IsInside(pane.Handle,focus) && ClassName(focus)==childClass)return true;
            // A click in another window/control wins over any outstanding extraction.
            if(focus==IntPtr.Zero ? !allowEmptyFocus : !IsInside(pane.Handle,focus))return false;
            var target=IntPtr.Zero;
            EnumWindowProc find=(handle,unused)=>{
                if(!IsInside(pane.Handle,handle) && ClassName(handle)==childClass && IsWindowVisible(handle) && IsWindowEnabled(handle))
                {target=handle;return false;}return true;
            };
            if(ClassName(window)==childClass && IsWindowVisible(window) && IsWindowEnabled(window))target=window;
            else EnumChildWindows(window,find,IntPtr.Zero);
            if(target==IntPtr.Zero || GetWindowThreadProcessId(target,IntPtr.Zero)!=GetCurrentThreadId())return false;
            SetFocus(target);
            return GetFocus()==target;
        }
        private static bool IsInside(IntPtr parent,IntPtr child)=>parent==child || IsChild(parent,child);
        private static string ClassName(IntPtr window)
        {var name=new StringBuilder(256);GetClassName(window,name,name.Capacity);return name.ToString();}
        private delegate bool EnumWindowProc(IntPtr window,IntPtr parameter);
        [DllImport("user32.dll")] private static extern IntPtr GetFocus();
        [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window,uint flags);
        [DllImport("user32.dll")][return:MarshalAs(UnmanagedType.Bool)] private static extern bool IsChild(IntPtr parent,IntPtr child);
        [DllImport("user32.dll")][return:MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowEnabled(IntPtr window);
        [DllImport("user32.dll")][return:MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr window,StringBuilder name,int maximum);
        [DllImport("user32.dll")][return:MarshalAs(UnmanagedType.Bool)] private static extern bool EnumChildWindows(IntPtr window,EnumWindowProc callback,IntPtr parameter);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window,IntPtr process);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    }
}
