using System;
using System.IO;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.Infrastructure
{
    internal sealed class WorkbookProjectContext
    {
        private string currentWorkbookPath;

        public ProjectStore Store { get; private set; }
        public ProjectState State { get; private set; }
        public DocumentImporter Importer { get; private set; }
        public SnipService Snips { get; private set; }
        public DocumentMatcher Matcher { get; private set; }

        public void Ensure(ExcelInterop.Workbook workbook)
        {
            if (workbook == null) throw new InvalidOperationException("Open an Excel workbook first.");
            if (string.IsNullOrWhiteSpace(workbook.Path))
                throw new InvalidOperationException("Save the workbook before creating its Doctracker file.");

            var workbookPath = Path.GetFullPath(workbook.FullName);
            if (string.Equals(workbookPath, currentWorkbookPath, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var directoryName = "." + Path.GetFileNameWithoutExtension(workbook.Name) + ".doctracker";
            var projectDirectory = Path.Combine(workbook.Path, directoryName);
            Store = new ProjectStore(projectDirectory);
            State = Store.LoadOrCreate(workbookPath);
            State.WorkbookPath = workbookPath;
            Importer = new DocumentImporter(Store);
            Snips = new SnipService(Store, new TextValueParser());
            Matcher = new DocumentMatcher();
            Store.Save(State);
            currentWorkbookPath = workbookPath;
        }
    }
}
