using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public interface IWorkbookParts
    {
        IEnumerable<string> Ids(string xmlNamespace);
        string Read(string id);
        string Add(string xml);
        void Delete(string id);
    }
    public sealed class PortableFile
    {
        public string Path { get; set; }
        public string Hash { get; set; }
        public long Length { get; set; }
        public List<string> Parts { get; set; } = new List<string>();
    }
    public sealed class PortableManifest
    {
        public int Version { get; set; } = 1;
        public long Generation { get; set; }
        public string Metadata { get; set; }
        public List<PortableFile> Files { get; set; } = new List<PortableFile>();
    }
    // Chunked immutable attachments. Only new files and the small manifest cross COM on save.
    public sealed class PortableProject
    {
        public const string ManifestNamespace = "urn:doctracker:portable:1";
        public const string BlobNamespace = "urn:doctracker:blob:1";
        public const long MaximumEmbeddedBytes = 256L * 1024 * 1024;
        private const int ChunkSize = 1024 * 1024;
        private readonly IWorkbookParts parts;
        private string currentId;
        private PortableManifest current;
        private static readonly XmlSerializer serializer = new XmlSerializer(typeof(PortableManifest));
        public PortableProject(IWorkbookParts parts) { this.parts=parts; ReadCurrent(); }
        public bool Exists => current != null;
        private void ReadCurrent()
        {
            var manifests=parts.Ids(ManifestNamespace).Select(id => new {Id=id,Manifest=DecodeManifest(parts.Read(id))}).OrderByDescending(x=>x.Manifest.Generation).ToList();
            if(manifests.Count>0) { currentId=manifests[0].Id;current=manifests[0].Manifest; }
        }
        public ProjectState Restore(ProjectStore store)
        {
            if(current==null) return null;
            ProjectState state;
            using(var input=new MemoryStream(Unpack(current.Metadata))) state=ProjectStore.ReadMetadata(input);
            ProjectStore.AtomicWrite(store.MetadataPath,ProjectStore.MetadataBytes(state));
            var files = current.Files.ToDictionary(f => f.Path, StringComparer.Ordinal);
            store.IndexResolver = key => {
                if (files.TryGetValue("indexes/" + key + ".xml.gz", out var file)) Extract(file, store.ProjectDirectory);
            };
            // PDFs/images are materialized only when requested by the viewer, OCR or export.
            store.DocumentResolver = document => {
                files.TryGetValue(document.RelativePath.Replace('\\','/'), out var file);
                if(file!=null) Extract(file,store.ProjectDirectory);
            };
            return store.LoadOrCreate(state.WorkbookPath);
        }
        public void Save(ProjectStore store, ProjectState state)
        {
            var observed=new PortableProject(parts);
            if(observed.currentId!=currentId) throw new IOException("Les preuves du classeur ont changé dans une autre session. Rechargez le classeur avant d'enregistrer pour éviter un écrasement.");
            var previous=(current?.Files??new List<PortableFile>()).ToDictionary(f=>f.Path,StringComparer.Ordinal);
            Func<string,string,PortableFile> unchanged=(path,hash)=>previous.TryGetValue(path,out var file) && file.Hash==hash ? file : null;
            var inputs=new List<Tuple<string,string,string>>();
            if(string.IsNullOrEmpty(state.SharedVaultPath))
                foreach(var doc in state.Documents) inputs.Add(Tuple.Create(doc.RelativePath.Replace('\\','/'),unchanged(doc.RelativePath.Replace('\\','/'),doc.Sha256)!=null ? store.LocalDocumentPath(doc) : store.ResolveDocumentPath(doc),doc.Sha256));
            foreach(var doc in state.Documents.Where(d=>!string.IsNullOrEmpty(d.IndexKey)))
            {
                var relative = "indexes/" + doc.IndexKey + ".xml.gz";
                if (previous.TryGetValue(relative, out var existingIndex))
                    inputs.Add(Tuple.Create(relative, store.IndexPath(doc.IndexKey), existingIndex.Hash));
                else
                {
                    var path = store.ResolveIndexPath(doc.IndexKey);
                    if (File.Exists(path)) inputs.Add(Tuple.Create(relative, path, DocumentImporter.ComputeSha256(path)));
                }
            }
            if(inputs.Sum(x=>unchanged(x.Item1,x.Item3)?.Length ?? new FileInfo(x.Item2).Length)>MaximumEmbeddedBytes) throw new IOException("Les pièces dépassent 256 Mo. Activez le mode Documents partagés pour garder le classeur léger, ou réduisez le dossier.");
            var next=new PortableManifest {Generation=(current?.Generation??0)+1,Metadata=Pack(ProjectStore.MetadataBytes(state))};
            var added=new List<string>(); string manifestId=null;
            try
            {
                foreach(var input in inputs)
                {
                    var existing=unchanged(input.Item1,input.Item3);
                    if(existing!=null) { next.Files.Add(existing);continue; }
                    var file=new PortableFile {Path=input.Item1,Hash=input.Item3,Length=new FileInfo(input.Item2).Length};
                    if(file.Path.StartsWith("documents/",StringComparison.Ordinal) && DocumentImporter.ComputeSha256(input.Item2)!=file.Hash) throw new InvalidDataException("Pièce modifiée : "+file.Path);
                    using(var stream=File.OpenRead(input.Item2))
                    {
                        var buffer=new byte[ChunkSize];int count;
                        while((count=stream.Read(buffer,0,buffer.Length))>0)
                        { var id=parts.Add("<blob xmlns=\""+BlobNamespace+"\">"+Convert.ToBase64String(buffer,0,count)+"</blob>");added.Add(id);file.Parts.Add(id); }
                    }
                    next.Files.Add(file);
                }
                using(var data=new MemoryStream()) { serializer.Serialize(data,next);manifestId=parts.Add("<project xmlns=\""+ManifestNamespace+"\">"+Pack(data.ToArray())+"</project>"); }
            }
            catch { foreach(var id in added) try {parts.Delete(id);} catch {} throw; }
            // New manifest is complete before removing anything used by the previous version.
            currentId=manifestId;current=next;
            try
            {
                // Remove all obsolete manifests first. If cleanup fails, keep their blobs too.
                foreach (var id in parts.Ids(ManifestNamespace).Where(id => id != manifestId).ToList()) parts.Delete(id);
                var used=new HashSet<string>(next.Files.SelectMany(f=>f.Parts));
                foreach(var id in parts.Ids(BlobNamespace).Where(id=>!used.Contains(id)).ToList()) parts.Delete(id);
            }
            catch (Exception failure) { System.Diagnostics.Trace.WriteLine("Doctracker attachment cleanup deferred: " + failure); }
        }
        private void Extract(PortableFile file,string root)
        {
            var path=SafePath(root,file.Path); if(File.Exists(path) && DocumentImporter.ComputeSha256(path)==file.Hash) return;
            if(file.Length<0 || file.Length>MaximumEmbeddedBytes) throw new InvalidDataException("Pièce intégrée trop volumineuse.");
            Directory.CreateDirectory(Path.GetDirectoryName(path));var temp=path+".tmp-"+Guid.NewGuid().ToString("N");
            try
            {
                long length=0;
                using(var stream=File.Create(temp)) foreach(var id in file.Parts)
                { var bytes=Convert.FromBase64String(ReadXml(parts.Read(id)).InnerText);length+=bytes.Length;if(length>file.Length)throw new InvalidDataException("Longueur incorrecte.");stream.Write(bytes,0,bytes.Length); }
                if(length!=file.Length || DocumentImporter.ComputeSha256(temp)!=file.Hash) throw new InvalidDataException("Pièce intégrée corrompue : "+file.Path);
                if(File.Exists(path)) File.Delete(path); File.Move(temp,path);
            }
            finally { if(File.Exists(temp))File.Delete(temp); }
        }
        private static PortableManifest DecodeManifest(string xml)
        {
            using(var input=new MemoryStream(Unpack(ReadXml(xml).InnerText)))
            using(var reader=XmlReader.Create(input,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))
            { var manifest=(PortableManifest)serializer.Deserialize(reader); if(manifest.Version!=1)throw new InvalidDataException("Version de dossier inconnue.");return manifest; }
        }
        private static XmlDocument ReadXml(string xml)
        {
            var doc=new XmlDocument {XmlResolver=null};using(var input=XmlReader.Create(new StringReader(xml),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=400L*1024*1024}))doc.Load(input);return doc;
        }
        public static string Pack(byte[] data)
        { using(var output=new MemoryStream()){ using(var zip=new GZipStream(output,CompressionMode.Compress,true))zip.Write(data,0,data.Length);return Convert.ToBase64String(output.ToArray());} }
        public static byte[] Unpack(string data)
        {
            using(var input=new MemoryStream(Convert.FromBase64String(data)))using(var zip=new GZipStream(input,CompressionMode.Decompress))using(var output=new MemoryStream())
            { var buffer=new byte[81920];int n;while((n=zip.Read(buffer,0,buffer.Length))>0){if(output.Length+n>MaximumEmbeddedBytes)throw new InvalidDataException("Métadonnées trop volumineuses.");output.Write(buffer,0,n);}return output.ToArray(); }
        }
        public static string SafePath(string root,string relative)
        {
            if(string.IsNullOrWhiteSpace(relative)||relative.Contains(":")||relative.StartsWith("/")||relative.StartsWith("\\"))throw new InvalidDataException("Chemin de sauvegarde incorrect.");
            var path=System.IO.Path.GetFullPath(System.IO.Path.Combine(root,relative.Replace('\\',System.IO.Path.DirectorySeparatorChar).Replace('/',System.IO.Path.DirectorySeparatorChar)));
            if(!path.StartsWith(System.IO.Path.GetFullPath(root).TrimEnd(System.IO.Path.DirectorySeparatorChar)+System.IO.Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("Chemin hors du dossier.");return path;
        }
    }
}
