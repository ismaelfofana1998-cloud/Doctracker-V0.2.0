using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Doctracker.Core.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Doctracker.AddIn.Infrastructure
{
    internal sealed class ExcelWorkbookParts : IWorkbookParts, IDisposable
    {
        private readonly ExcelInterop.Workbook workbook;
        private readonly Control dispatcher = new Control();
        public ExcelWorkbookParts(ExcelInterop.Workbook workbook) { this.workbook=workbook;dispatcher.CreateControl(); }
        private T OnExcel<T>(Func<T> action) => dispatcher.InvokeRequired ? (T)dispatcher.Invoke(action) : action();
        public IEnumerable<string> Ids(string ns) => OnExcel(() => {
            var result=new List<string>();var selected=workbook.CustomXMLParts.SelectByNamespace(ns);
            for(var i=1;i<=selected.Count;i++)result.Add(selected[i].Id);return result;
        });
        public string Read(string id) => OnExcel(() => workbook.CustomXMLParts.SelectByID(id)?.XML ?? throw new InvalidOperationException("Pièce intégrée manquante. Utilisez une sauvegarde Doctracker."));
        public string Add(string xml) => OnExcel(() => workbook.CustomXMLParts.Add(xml,Type.Missing).Id);
        public void Delete(string id) => OnExcel(() => {workbook.CustomXMLParts.SelectByID(id)?.Delete();return true;});
        public void Dispose() {dispatcher.Dispose();}
    }
}
