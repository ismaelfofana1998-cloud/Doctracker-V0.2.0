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
            Application.SheetBeforeDoubleClick += Application_SheetBeforeDoubleClick;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.SheetBeforeDoubleClick -= Application_SheetBeforeDoubleClick;
            if (Controller != null) Controller.Dispose();
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
