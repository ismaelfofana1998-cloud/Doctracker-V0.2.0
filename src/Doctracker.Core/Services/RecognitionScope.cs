using System.Collections.Generic;
using System.Linq;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public static class RecognitionScope
    {
        // Null means all folders; empty means unfiled. An explicit ID selection is
        // intersected with the folder, never expanded by a document's other categories.
        public static List<DocumentRecord> Select(ProjectState state,string folder,IEnumerable<string> selectedIds=null)
        {
            var selected=selectedIds==null?null:new HashSet<string>(selectedIds);
            return state.Documents.Where(d=>(selected==null || selected.Contains(d.Id)) &&
                (folder==null || (folder.Length==0?d.Categories.Count==0:d.Categories.Any(c=>ProjectFolders.Within(c,folder))))).ToList();
        }
    }
}
