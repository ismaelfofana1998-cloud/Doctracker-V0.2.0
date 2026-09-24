using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public static class RecoveryArchive
    {
        public static void Export(ProjectStore store, ProjectState state, string destination, CancellationToken cancellation)
        {
            var temp=destination+"."+Guid.NewGuid().ToString("N")+".tmp";
            try
            {
                using(var zip=ZipFile.Open(temp,ZipArchiveMode.Create))
                {
                    var entry=zip.CreateEntry("project.xml");using(var output=entry.Open()){var data=ProjectStore.MetadataBytes(state);output.Write(data,0,data.Length);}
                    foreach(var doc in state.Documents)
                    {
                        cancellation.ThrowIfCancellationRequested();var path=store.ResolveDocumentPath(doc);ProjectStore.VerifyHash(doc,path);
                        zip.CreateEntryFromFile(path,doc.RelativePath.Replace('\\','/'),CompressionLevel.NoCompression);
                        if(!string.IsNullOrEmpty(doc.IndexKey) && File.Exists(store.ResolveIndexPath(doc.IndexKey)))zip.CreateEntryFromFile(store.ResolveIndexPath(doc.IndexKey),"indexes/"+doc.IndexKey+".xml.gz",CompressionLevel.NoCompression);
                    }
                }
                cancellation.ThrowIfCancellationRequested();
                if(File.Exists(destination)) File.Replace(temp,destination,destination+".bak");else File.Move(temp,destination);
            }
            finally {if(File.Exists(temp))File.Delete(temp);}
        }
        // Restore into a NEW directory; validation completes before the caller switches context.
        public static ProjectState Restore(string archive, ProjectStore destination, CancellationToken cancellation)
        {
            if(Directory.Exists(destination.ProjectDirectory) && Directory.EnumerateFileSystemEntries(destination.ProjectDirectory).Any())throw new IOException("Le dossier de restauration doit être vide.");
            using(var zip=ZipFile.OpenRead(archive))
            {
                if(zip.Entries.Count>100000 || zip.Entries.Sum(e=>e.Length)>16L*1024*1024*1024)throw new InvalidDataException("Sauvegarde trop volumineuse.");
                foreach(var entry in zip.Entries)
                {
                    cancellation.ThrowIfCancellationRequested();var path=PortableProject.SafePath(destination.ProjectDirectory,entry.FullName);
                    if(string.IsNullOrEmpty(entry.Name))continue;
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                    using(var input=entry.Open())using(var output=new FileStream(path,FileMode.CreateNew,FileAccess.Write))
                    { var buffer=new byte[81920];long length=0;int count;while((count=input.Read(buffer,0,buffer.Length))>0){cancellation.ThrowIfCancellationRequested();length+=count;if(length>entry.Length)throw new InvalidDataException("Longueur de sauvegarde incohérente.");output.Write(buffer,0,count);}if(length!=entry.Length)throw new InvalidDataException("Sauvegarde tronquée."); }
                }
            }
            if(!File.Exists(destination.MetadataPath))throw new InvalidDataException("Sauvegarde sans métadonnées.");
            var state=destination.LoadOrCreate(null);
            foreach(var doc in state.Documents){cancellation.ThrowIfCancellationRequested();ProjectStore.VerifyHash(doc,destination.LocalDocumentPath(doc));}
            foreach(var snip in state.Snips)if(!state.Documents.Any(doc=>doc.Id==snip.DocumentId))throw new InvalidDataException("Preuve sans document source.");
            return state;
        }
    }
}
