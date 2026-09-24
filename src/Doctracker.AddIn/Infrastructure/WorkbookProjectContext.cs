using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.Infrastructure
{
    internal sealed class WorkbookProjectContext : IDisposable
    {
        private string currentWorkbookPath;
        private ExcelInterop.Workbook workbook;
        private ExcelWorkbookParts parts;
        private PortableProject portable;
        private volatile bool dirty;
        public bool IsBusy { get; set; }
        public string WorkbookPath => currentWorkbookPath;
        public ProjectStore Store { get; private set; }
        public ProjectState State { get; private set; }
        public DocumentImporter Importer { get; private set; }
        public SnipService Snips { get; private set; }
        public DocumentMatcher Matcher { get; private set; }
        public static string CacheRoot => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"Doctracker","Recovery");
        private static bool Supported(ExcelInterop.Workbook book) => new[]{".xlsx",".xlsm",".xlsb"}.Contains(Path.GetExtension(book.Name).ToLowerInvariant());
        public void Ensure(ExcelInterop.Workbook book)
        {
            if(book==null)throw new InvalidOperationException("Ouvrez un classeur Excel.");
            if(string.IsNullOrWhiteSpace(book.Path)||!Supported(book))throw new InvalidOperationException("Enregistrez le classeur au format .xlsx, .xlsm ou .xlsb pour conserver ses pièces intégrées.");
            if(State!=null) {currentWorkbookPath=book.FullName;State.WorkbookPath=book.FullName;return;}
            workbook=book;parts=new ExcelWorkbookParts(book);portable=new PortableProject(parts);
            var store=new ProjectStore(Path.Combine(CacheRoot,Guid.NewGuid().ToString("N")));
            ProjectState state;
            if(portable.Exists) state=portable.Restore(store);
            else
            {
                var legacy=Path.Combine(book.Path,"."+Path.GetFileNameWithoutExtension(book.Name)+".doctracker");
                if(Directory.Exists(legacy)) {CopyDirectory(legacy,store.ProjectDirectory);state=store.LoadOrCreate(book.FullName);}
                else state=new ProjectState {WorkbookPath=book.FullName};
                dirty=true;
            }
            Bind(store,state);currentWorkbookPath=book.FullName;
            Store.Save(State);MarkWorkbookDirty();
        }
        private void Bind(ProjectStore store,ProjectState state)
        {
            Store=store;State=state;Store.SharedVaultPath=state.SharedVaultPath;
            Store.Saved+=()=>dirty=true;
            Importer=new DocumentImporter(store);Snips=new SnipService(store,new TextValueParser());Matcher=new DocumentMatcher();
        }
        public void MarkWorkbookDirty() { if(dirty && workbook!=null && !workbook.ReadOnly)workbook.Saved=false; }
        public void BeforeSave(bool saveAs)
        {
            if(State==null)return;
            if(IsBusy)throw new InvalidOperationException("Attendez ou annulez l'opération avant d'enregistrer.");
            if((workbook.ReadOnly&&!saveAs)||!Supported(workbook))throw new InvalidOperationException("Enregistrez une copie modifiable au format .xlsx, .xlsm ou .xlsb.");
            CaptureCellLinks();
            if(!dirty)return;
            if(!string.IsNullOrEmpty(State.SharedVaultPath)) SharedVault.Publish(State.SharedVaultPath,Store,State);
            portable.Save(Store,State);
            dirty=false;
        }
        private void CaptureCellLinks()
        {
            var links=new List<CellLinkRecord>();var known=new HashSet<string>(State.Snips.Select(s=>s.Id));
            var cells=new Excel.ExcelCellGateway(workbook.Application);
            foreach(ExcelInterop.Worksheet sheet in workbook.Worksheets)
            {
                ExcelInterop.Range comments;
                try{comments=sheet.Cells.SpecialCells(ExcelInterop.XlCellType.xlCellTypeComments);}
                catch(System.Runtime.InteropServices.COMException){continue;}
                foreach(ExcelInterop.Range cell in comments.Cells)
                {
                    var ids=cells.GetSnipIds(cell).Where(known.Contains).ToList();if(ids.Count==0)continue;
                    links.Add(new CellLinkRecord {WorksheetName=sheet.Name,CellAddress=cell.Address[false,false,ExcelInterop.XlReferenceStyle.xlA1],SnipIds=ids});
                }
            }
            Func<CellLinkRecord,string> key=x=>x.WorksheetName+"!"+x.CellAddress+"|"+string.Join(",",x.SnipIds);
            if(!links.Select(key).OrderBy(x=>x).SequenceEqual(State.CellLinks.Select(key).OrderBy(x=>x)))
            {
                State.CellLinks=links;
                var byId=State.Snips.ToDictionary(s=>s.Id);var seen=new HashSet<string>();
                foreach(var link in links)foreach(var id in link.SnipIds)if(seen.Add(id)){byId[id].WorksheetName=link.WorksheetName;byId[id].CellAddress=link.CellAddress;}
                Store.Save(State);
            }
        }
        public void RestoreMetadata(string file)
        {
            ProjectState state;using(var input=File.OpenRead(file))state=ProjectStore.ReadMetadata(input);
            var source=Path.GetDirectoryName(file);
            if(Path.GetFileName(source)=="recovery")source=Path.GetDirectoryName(source);
            var store=new ProjectStore(Path.Combine(CacheRoot,Guid.NewGuid().ToString("N")));
            foreach(var folder in new[]{"documents","indexes"})if(Directory.Exists(Path.Combine(source,folder)))CopyDirectory(Path.Combine(source,folder),Path.Combine(store.ProjectDirectory,folder));
            ProjectStore.AtomicWrite(store.MetadataPath,ProjectStore.MetadataBytes(state));
            Bind(store,store.LoadOrCreate(workbook.FullName));Store.Save(State);MarkWorkbookDirty();
        }
        public void AfterSave(bool success) {if(!success)dirty=true;}
        public void RestoreArchive(string archive)
        {
            var store=new ProjectStore(Path.Combine(CacheRoot,Guid.NewGuid().ToString("N")));
            var restored=RecoveryArchive.Restore(archive,store,System.Threading.CancellationToken.None);
            restored.SharedVaultPath="";restored.WorkbookPath=workbook.FullName;
            Bind(store,restored);Store.Save(restored);MarkWorkbookDirty();
        }
        private static void CopyDirectory(string source,string destination)
        {
            Directory.CreateDirectory(destination);
            foreach(var file in Directory.GetFiles(source))File.Copy(file,Path.Combine(destination,Path.GetFileName(file)),false);
            foreach(var dir in Directory.GetDirectories(source))if((File.GetAttributes(dir)&FileAttributes.ReparsePoint)==0)CopyDirectory(dir,Path.Combine(destination,Path.GetFileName(dir)));
        }
        public void Dispose(){parts?.Dispose();}
    }
}
