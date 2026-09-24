using System;
using Microsoft.Office.Core;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Doctracker.AddIn.Ribbon;
using Doctracker.AddIn.UI;

namespace Doctracker.AddIn
{
    public partial class ThisAddIn
    {
        internal PaneController Controller { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Controller = new PaneController(this, Application);
            Application.WindowActivate += Application_WindowActivate;
            Application.WorkbookAfterSave += Application_WorkbookAfterSave;
            Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
            Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            Application.SheetSelectionChange += Application_SheetSelectionChange;
            Application.SheetBeforeDoubleClick += Application_SheetBeforeDoubleClick;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.WindowActivate -= Application_WindowActivate;
            Application.WorkbookAfterSave -= Application_WorkbookAfterSave;
            Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
            Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            Application.SheetSelectionChange -= Application_SheetSelectionChange;
            Application.SheetBeforeDoubleClick -= Application_SheetBeforeDoubleClick;
            if (Controller != null) Controller.Dispose();
        }

        private void Application_SheetSelectionChange(object sheet, ExcelInterop.Range target)
        {
            Controller?.SelectionChanged(target);
        }

        private void Application_SheetBeforeDoubleClick(
            object sheet,
            ExcelInterop.Range target,
            ref bool cancel)
        {
            if (Controller != null && Controller.TryNavigateFromCell(target))
            {
                cancel = true;
            }
        }

        private void Application_WindowActivate(ExcelInterop.Workbook workbook, ExcelInterop.Window window)
        {
            try { Controller?.RefreshVisible(); DoctrackerRibbon.Instance?.Refresh(); }
            catch (System.Runtime.InteropServices.COMException exception) { System.Diagnostics.Trace.WriteLine(exception); }
        }
        private void Application_WorkbookAfterSave(ExcelInterop.Workbook workbook, bool success)
        {
            Controller?.AfterSave(workbook,success);
            // Saving does not change document selection or require rebinding the pane.
        }
        private void Application_WorkbookBeforeSave(ExcelInterop.Workbook workbook, bool saveAs, ref bool cancel)
        {
            if (Controller?.IsBusy(workbook) == true)
            {
                cancel = true;
                System.Windows.Forms.MessageBox.Show("Attendez ou annulez l'opération Doctracker avant d'enregistrer.", "Doctracker");
                return;
            }
            try { Controller?.BeforeSave(workbook,saveAs); }
            catch(Exception exception) { cancel=true;System.Windows.Forms.MessageBox.Show("Enregistrement interrompu pour conserver les preuves : " + exception.Message,"Doctracker"); }
        }
        private void Application_WorkbookBeforeClose(ExcelInterop.Workbook workbook, ref bool cancel)
        {
            if (Controller?.IsBusy(workbook) == true)
            {
                cancel = true;
                System.Windows.Forms.MessageBox.Show("Attendez ou annulez l'opération Doctracker avant de fermer ce classeur.", "Doctracker");
            }
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new DoctrackerRibbon();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        #endregion
    }
}
