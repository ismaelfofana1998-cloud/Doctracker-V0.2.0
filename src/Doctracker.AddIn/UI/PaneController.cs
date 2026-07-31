using System;
using Doctracker.Core.Models;
using Microsoft.Office.Core;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.UI
{
    internal sealed class PaneController : IDisposable
    {
        private readonly DoctrackerPaneControl control;
        private readonly Microsoft.Office.Tools.CustomTaskPane pane;

        public PaneController(ThisAddIn addIn, ExcelInterop.Application application)
        {
            if (addIn == null) throw new ArgumentNullException(nameof(addIn));
            control = new DoctrackerPaneControl(application);
            pane = addIn.CustomTaskPanes.Add(control, "Doctracker");
            pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            pane.Width = 760;
            pane.Visible = false;
        }

        public void Toggle()
        {
            pane.Visible = !pane.Visible;
            if (pane.Visible) control.RefreshProject();
        }

        public void Show()
        {
            pane.Visible = true;
            control.RefreshProject();
        }

        public void ImportDocuments()
        {
            Show();
            control.ImportDocuments();
        }

        public void CreateSnip(SnipType type)
        {
            Show();
            control.SetSnipMode(type);
        }

        public void SetSnipMode(SnipType? type)
        {
            Show();
            control.SetSnipMode(type);
        }

        public bool IsSnipMode(SnipType type)
        {
            return control.IsSnipMode(type);
        }

        public void MatchSelection()
        {
            Show();
            control.MatchSelection();
        }

        public void SetMatchingInputSelection()
        {
            Show();
            control.SetMatchingInputSelection();
        }

        public void SetMatchingOutputSelection()
        {
            Show();
            control.SetMatchingOutputSelection();
        }

        public void SearchSelection()
        {
            Show();
            control.SearchSelection();
        }

        public void NavigateFromSelection()
        {
            Show();
            control.NavigateFromSelection();
        }

        public void ReviewSelection()
        {
            Show();
            control.ReviewSelection();
        }

        public bool TryNavigateFromCell(ExcelInterop.Range target)
        {
            var result = control.TryNavigateFromCell(target);
            if (result) pane.Visible = true;
            return result;
        }

        public void Dispose()
        {
            control.Dispose();
        }
    }
}
