using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Doctracker.AddIn.Infrastructure;
using Doctracker.Core.Models;
using Microsoft.Office.Core;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.UI
{
    internal sealed class PaneController : IDisposable
    {
        private readonly ThisAddIn addIn;
        private readonly ExcelInterop.Application application;
        private readonly Dictionary<int, WindowPane> panes = new Dictionary<int, WindowPane>();
        private readonly Dictionary<ExcelInterop.Workbook, WorkbookProjectContext> contexts = new Dictionary<ExcelInterop.Workbook, WorkbookProjectContext>();
        public PaneController(ThisAddIn addIn, ExcelInterop.Application application) { this.addIn = addIn; this.application = application; }

        private WindowPane Current()
        {
            CleanupClosedWindows();
            var workbook = application.ActiveWorkbook;
            var window = application.ActiveWindow;
            if (workbook == null || window == null) throw new InvalidOperationException("Ouvrez un classeur Excel.");
            WindowPane entry;
            if (panes.TryGetValue(window.Hwnd, out entry))
            {
                if (Equals(entry.Workbook, workbook)) return entry;
                addIn.CustomTaskPanes.Remove(entry.Pane); entry.Control.Dispose(); panes.Remove(window.Hwnd);
            }
            WorkbookProjectContext context;
            if (!contexts.TryGetValue(workbook, out context)) contexts[workbook] = context = new WorkbookProjectContext();
            var control = new DoctrackerPaneControl(application, workbook, context);
            var pane = addIn.CustomTaskPanes.Add(control, "Doctracker", window);
            pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            pane.Width = Math.Max(400, Math.Min(760, (int)(window.Width * 0.55)));
            entry = new WindowPane { Workbook = workbook, Control = control, Pane = pane };
            panes[window.Hwnd] = entry;
            return entry;
        }

        private void Run(Action<WindowPane> action)
        {
            try { action(Current()); }
            catch (Exception exception) { MessageBox.Show(exception.Message, "Doctracker", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        private void Show(Action<DoctrackerPaneControl> action) => Run(entry =>
        { entry.Pane.Visible = true; entry.Control.RefreshProject(); action(entry.Control); });
        public void Toggle() => Run(entry => { entry.Pane.Visible = !entry.Pane.Visible; if (entry.Pane.Visible) entry.Control.RefreshProject(); });
        public void ImportDocuments() => Show(control => control.ImportDocuments());
        public void SetSnipMode(SnipType? type) => Show(control => control.SetSnipMode(type));
        public bool IsSnipMode(SnipType type)
        {
            var window = application.ActiveWindow;
            WindowPane entry;
            return window != null && panes.TryGetValue(window.Hwnd, out entry) && entry.Control.IsSnipMode(type);
        }
        public void MatchSelection() => Show(control => control.MatchSelection());
        public void SetMatchingInputSelection() => Show(control => control.SetMatchingInputSelection());
        public void SetMatchingOutputSelection() => Show(control => control.SetMatchingOutputSelection());
        public void SearchSelection() => Show(control => control.SearchSelection());
        public void NavigateFromSelection() => Show(control => control.NavigateFromSelection());
        public void ReviewSelection() => Show(control => control.ReviewSelection());
        public bool TryNavigateFromCell(ExcelInterop.Range target)
        {
            // Normal Excel double-click must never create a project or show an error.
            if (target == null || target.Cells.CountLarge != 1 || target.Comment == null ||
                !((string)target.Comment.Text()).Contains(Excel.ExcelCellGateway.MarkerPrefix)) return false;
            var found = false;
            Run(entry => { found = entry.Control.TryNavigateFromCell(target); if (found) entry.Pane.Visible = true; });
            return found;
        }
        public void SelectionChanged(ExcelInterop.Range target)
        {
            // Excel fires this for mouse clicks AND keyboard navigation. Never steal focus.
            try
            {
                var book = application.ActiveWorkbook;
                if (book == null || IsBusy(book)) return;
                var window = application.ActiveWindow;
                WindowPane existing;
                if (target == null || target.Cells.CountLarge != 1 || target.Comment == null ||
                    !((string)target.Comment.Text()).Contains(Excel.ExcelCellGateway.MarkerPrefix))
                {
                    if (window != null && panes.TryGetValue(window.Hwnd, out existing)) existing.Control.ClearCellProof();
                    return;
                }
                var entry = Current();
                if (entry.Control.TryNavigateFromCell(target)) entry.Pane.Visible = true;
            }
            catch (System.Runtime.InteropServices.COMException) { /* Excel can be editing or closing. */ }
            catch (Exception exception) { System.Diagnostics.Trace.WriteLine("Doctracker selection: " + exception.Message); }
        }

        public void BeforeSave(ExcelInterop.Workbook workbook,bool saveAs) { if(contexts.TryGetValue(workbook,out var context))context.BeforeSave(saveAs); }
        public void AfterSave(ExcelInterop.Workbook workbook,bool success) { if(contexts.TryGetValue(workbook,out var context))context.AfterSave(success); }
        public bool IsBusy(ExcelInterop.Workbook workbook) => contexts.TryGetValue(workbook, out var context) && context.IsBusy;
        public void RefreshVisible()
        {
            CleanupClosedWindows();
            var window = application.ActiveWindow;
            if (window != null && panes.TryGetValue(window.Hwnd, out var entry) && entry.Pane.Visible) entry.Control.RefreshProject();
        }
        private void CleanupClosedWindows()
        {
            var openWindows = new HashSet<int>();
            var openBooks = new HashSet<ExcelInterop.Workbook>();
            foreach (ExcelInterop.Workbook book in application.Workbooks)
            {
                openBooks.Add(book);
                foreach (ExcelInterop.Window window in book.Windows) openWindows.Add(window.Hwnd);
            }
            foreach (var pair in panes.Where(pair => !openWindows.Contains(pair.Key) || !openBooks.Contains(pair.Value.Workbook)).ToList())
            {
                // VSTO may already have removed the task pane after its window closed.
                try { addIn.CustomTaskPanes.Remove(pair.Value.Pane); }
                catch (System.Runtime.InteropServices.COMException) { }
                catch (ArgumentException) { }
                pair.Value.Control.Dispose();
                panes.Remove(pair.Key);
            }
            foreach (var book in contexts.Keys.Where(book => !openBooks.Contains(book)).ToList()) { contexts[book].Dispose(); contexts.Remove(book); }
        }

        public void Dispose()
        {
            foreach (var entry in panes.Values) entry.Control.Dispose();
            foreach(var context in contexts.Values)context.Dispose();
            panes.Clear(); contexts.Clear();
        }
        private sealed class WindowPane
        {
            public ExcelInterop.Workbook Workbook;
            public DoctrackerPaneControl Control;
            public Microsoft.Office.Tools.CustomTaskPane Pane;
        }
    }
}
