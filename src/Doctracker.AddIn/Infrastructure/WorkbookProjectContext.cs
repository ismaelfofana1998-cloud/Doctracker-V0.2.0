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
        public bool IsBusy { get; set; }
        public string WorkbookPath => currentWorkbookPath;
        public ProjectStore Store { get; private set; }
        public ProjectState State { get; private set; }
        public DocumentImporter Importer { get; private set; }
        public SnipService Snips { get; private set; }
        public DocumentMatcher Matcher { get; private set; }

        public void Ensure(ExcelInterop.Workbook workbook)
        {
            if (workbook == null) throw new InvalidOperationException("Ouvrez un classeur Excel.");
            if (string.IsNullOrWhiteSpace(workbook.Path)) throw new InvalidOperationException("Enregistrez le classeur avant d'ajouter des pièces.");
            if (workbook.ReadOnly) throw new InvalidOperationException("Le classeur est en lecture seule. Enregistrez une copie modifiable.");
            Uri uri;
            if (Uri.TryCreate(workbook.FullName, UriKind.Absolute, out uri) && !uri.IsFile)
                throw new InvalidOperationException("Utilisez un classeur dans un dossier local ou synchronisé sur cet ordinateur.");
            var path = Path.GetFullPath(workbook.FullName);
            if (string.Equals(path, currentWorkbookPath, StringComparison.OrdinalIgnoreCase)) return;
            if (IsBusy) throw new InvalidOperationException("Attendez la fin de l'opération avant de changer le nom du classeur.");
            // Each context is bound to one Workbook COM object by PaneController.
            // A changed path here is Save As, not a different workbook.
            var directory = Path.Combine(Path.GetDirectoryName(path), "." + Path.GetFileNameWithoutExtension(path) + ".doctracker");
            var nextStore = new ProjectStore(directory);
            ProjectState nextState;
            if (State != null && !string.Equals(directory, Store.ProjectDirectory, StringComparison.OrdinalIgnoreCase))
            {
                if (Directory.Exists(directory))
                    throw new IOException("Un dossier Doctracker existe déjà pour ce nom. Choisissez un autre nom pour conserver les preuves.");
                CopyDirectory(Store.ProjectDirectory, directory);
                nextState = nextStore.LoadOrCreate(path);
                nextState.AuditTrail.Add(new AuditEventRecord { Actor = Environment.UserName, Action = "WorkbookSavedAs",
                    EntityType = "Project", EntityId = nextState.ProjectId, Details = currentWorkbookPath + " → " + path });
            }
            else nextState = nextStore.LoadOrCreate(path);
            nextState.WorkbookPath = path;
            nextStore.Save(nextState);
            Store = nextStore; State = nextState;
            Importer = new DocumentImporter(Store);
            Snips = new SnipService(Store, new TextValueParser());
            Matcher = new DocumentMatcher();
            currentWorkbookPath = path;
        }

        private static void CopyDirectory(string source, string destination)
        {
            Directory.CreateDirectory(destination);
            foreach (var file in Directory.GetFiles(source))
                File.Copy(file, Path.Combine(destination, Path.GetFileName(file)), false);
            foreach (var directory in Directory.GetDirectories(source))
                CopyDirectory(directory, Path.Combine(destination, Path.GetFileName(directory)));
        }
    }
}
