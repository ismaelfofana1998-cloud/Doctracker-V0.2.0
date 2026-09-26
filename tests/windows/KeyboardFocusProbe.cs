using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

// Exercise the production Win32 focus code on Windows, without requiring Office.
public static class KeyboardFocusProbe
{
    public static string Run(Assembly assembly)
    {
        var type=assembly.GetType("Doctracker.AddIn.Infrastructure.WorksheetKeyboardFocus",true);
        var flags=BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic;
        var handoff=type.GetMethod("ReturnToChild",flags);
        var capture=type.GetMethod("CaptureChild",flags);
        var restore=type.GetMethod("RestoreAfterPaneRefresh",flags);
        using(var host=new Form())
        using(var pane=new Panel())
        using(var paneInput=new TextBox())
        using(var sheet=new TextBox())
        using(var editor=new TextBox())
        using(var other=new Form())
        {
            host.Text="Doctracker keyboard focus regression";
            pane.SetBounds(150,10,120,100);paneInput.SetBounds(0,0,100,24);
            sheet.SetBounds(10,10,120,24);editor.SetBounds(10,50,120,24);
            pane.Controls.Add(paneInput);host.Controls.Add(pane);
            host.Controls.Add(sheet);host.Controls.Add(editor);
            host.Show();host.Activate();SetForegroundWindow(host.Handle);Application.DoEvents();
            string childClass=ClassName(sheet.Handle);
            Func<bool,bool> giveBack=empty=>(bool)handoff.Invoke(null,new object[]{pane,host.Handle,childClass,empty});
            Func<IntPtr> remember=()=>(IntPtr)capture.Invoke(null,new object[]{host.Handle,childClass});
            Func<IntPtr,bool> recover=saved=>(bool)restore.Invoke(null,new object[]{pane,host.Handle,saved});

            paneInput.Focus();Require(GetFocus()==paneInput.Handle,"pane setup");
            Require(giveBack(false),"handoff before disabling PDF");
            Require(GetFocus()==sheet.Handle || GetFocus()==editor.Handle,"target must be outside pane");
            // Select the worksheet proxy explicitly; enumeration order is not guaranteed.
            sheet.Focus();pane.Enabled=false;
            TypeChar('a');Require(sheet.Text=="a","first character after snip");
            pane.Enabled=true;

            // Returning to a previous cell can rebind/show/rebuild the proof pane.
            sheet.Focus();IntPtr saved=remember();Require(saved==sheet.Handle,"capture worksheet");
            paneInput.Focus();Require(recover(saved),"restore after proof activation");
            TypeChar('b');Require(sheet.Text=="ab","first character after proof refresh");
            paneInput.Focus();SetFocus(IntPtr.Zero);
            Require(recover(saved),"restore when native focus was cleared");
            TypeChar('c');Require(sheet.Text=="abc","first character after empty focus");

            // Do not override an edit control, another cell/view or another app.
            editor.Focus();Require(!recover(saved),"edit focus must win");
            Require(GetFocus()==editor.Handle,"editor focus retained");
            Require(!giveBack(false) || GetFocus()==editor.Handle,"no delayed focus theft");
            TypeChar('x');Require(editor.Text=="x","editor receives typing");
            SetFocus(IntPtr.Zero);Require(!giveBack(false),"empty focus requires explicit handoff");
            paneInput.Focus();sheet.Enabled=false;
            Require(!recover(saved),"disabled worksheet excluded");sheet.Enabled=true;
            paneInput.Focus();sheet.Hide();Require(!recover(saved),"hidden worksheet excluded");sheet.Show();
            paneInput.Focus();Require((IntPtr)capture.Invoke(null,new object[]{host.Handle,"EXCEL7"})==IntPtr.Zero,"non-grid focus skips passive navigation");
            sheet.Focus();other.Show();other.Activate();SetForegroundWindow(other.Handle);Application.DoEvents();
            Require(GetForegroundWindow()==other.Handle,"other window setup");
            Require(remember()==IntPtr.Zero && !recover(saved) && !giveBack(true),"other foreground window respected");
            other.Close();host.Activate();SetForegroundWindow(host.Handle);
            sheet.Focus();saved=remember();sheet.Dispose();
            Require(!recover(saved),"destroyed worksheet excluded");
            host.Close();
        }
        return "PASS: keyboard focus handoff, first characters, passive proof refresh, editing/window/hidden/disposed guards";
    }
    private static void Require(bool value,string scenario){if(!value)throw new Exception("Keyboard focus: "+scenario);}
    private static void TypeChar(char value){SendMessage(GetFocus(),0x0102,new IntPtr(value),IntPtr.Zero);}
    private static string ClassName(IntPtr handle){var value=new StringBuilder(256);GetClassName(handle,value,value.Capacity);return value.ToString();}
    [DllImport("user32.dll")] private static extern IntPtr GetFocus();
    [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr window);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr window,StringBuilder value,int length);
    [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window,uint message,IntPtr wParam,IntPtr lParam);
}
