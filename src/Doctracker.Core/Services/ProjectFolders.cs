using System;
using System.Collections.Generic;
using System.Linq;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    // Logical workbook folders. Moving never renames a source file or changes a proof ID.
    public sealed class ProjectFolders
    {
        private readonly ProjectStore store;
        public ProjectFolders(ProjectStore store){this.store=store;}
        public static IEnumerable<string> Paths(ProjectState state)
        {
            var paths=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(var path in state.Folders.Concat(state.Documents.SelectMany(d=>d.Categories)))
            {
                var parts=path.Split(new[]{" / "},StringSplitOptions.RemoveEmptyEntries);
                for(var i=1;i<=parts.Length;i++)paths.Add(string.Join(" / ",parts.Take(i)));
            }
            return paths.OrderBy(p=>p,StringComparer.OrdinalIgnoreCase);
        }
        public static bool Within(string path,string parent)=>string.Equals(path,parent,StringComparison.OrdinalIgnoreCase)||path.StartsWith(parent+" / ",StringComparison.OrdinalIgnoreCase);
        public void Create(ProjectState state,string parent,string name)
        {
            var path=Join(parent,name);
            if(Paths(state).Contains(path,StringComparer.OrdinalIgnoreCase))throw new InvalidOperationException("Ce dossier existe déjà.");
            Change(state,()=>state.Folders.Add(path));
        }
        public void Move(ProjectState state,IEnumerable<string> documentIds,string destination)
        {
            if(!string.IsNullOrEmpty(destination)&&!Paths(state).Contains(destination,StringComparer.OrdinalIgnoreCase))throw new InvalidOperationException("Dossier introuvable.");
            var ids=new HashSet<string>(documentIds);var docs=state.Documents.Where(d=>ids.Contains(d.Id)).ToList();
            if(docs.Count!=ids.Count)throw new InvalidOperationException("Un document n'existe plus.");
            Change(state,()=>{foreach(var doc in docs)doc.Categories=string.IsNullOrEmpty(destination)?new List<string>():new List<string>{destination};});
        }
        public void Rename(ProjectState state,string path,string name)
        {
            var parent=Parent(path);var next=Join(parent,name);
            if(!string.Equals(path,next,StringComparison.OrdinalIgnoreCase)&&Paths(state).Contains(next,StringComparer.OrdinalIgnoreCase))throw new InvalidOperationException("Ce dossier existe déjà.");
            Change(state,()=>{
                state.Folders=Paths(state).Select(p=>Within(p,path)?next+p.Substring(path.Length):p).ToList();
                foreach(var doc in state.Documents)doc.Categories=doc.Categories.Select(p=>Within(p,path)?next+p.Substring(path.Length):p).ToList();
            });
        }
        public void Remove(ProjectState state,string path)
        {
            var parent=Parent(path);
            Change(state,()=>{
                state.Folders=Paths(state).Where(p=>!Within(p,path)).ToList();
                foreach(var doc in state.Documents)doc.Categories=doc.Categories.Select(p=>Within(p,path)?parent:p).Where(p=>p.Length>0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            });
        }
        public static string Parent(string path){var index=path.LastIndexOf(" / ",StringComparison.Ordinal);return index<0?"":path.Substring(0,index);}
        private static string Join(string parent,string name)
        {
            name=(name??"").Trim();if(name.Length==0||name.Length>80||name.IndexOfAny(new[]{'/','\\',';','\r','\n'})>=0)throw new ArgumentException("Nom de dossier : 1 à 80 caractères, sans /, \\, ; ni retour à la ligne.");
            return string.IsNullOrEmpty(parent)?name:parent+" / "+name;
        }
        private void Change(ProjectState state,Action change)
        {
            var folders=state.Folders.ToList();var categories=state.Documents.ToDictionary(d=>d.Id,d=>d.Categories.ToList());
            try{change();store.Save(state);}
            catch{state.Folders=folders;foreach(var doc in state.Documents)doc.Categories=categories[doc.Id];throw;}
        }
    }
}
