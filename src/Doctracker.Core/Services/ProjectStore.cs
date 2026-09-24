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
    public sealed class ProjectStore
    {
        private readonly object sync = new object();
        private readonly System.Collections.Concurrent.ConcurrentDictionary<string,string> verified = new System.Collections.Concurrent.ConcurrentDictionary<string,string>();
        private static readonly XmlSerializer reader = new XmlSerializer(typeof(ProjectState));
        private static readonly XmlSerializer writer = CreateMetadataSerializer();
        private static readonly XmlSerializer indexSerializer = new XmlSerializer(typeof(List<PageTextRecord>));
        public event Action Saved;
        public Action<DocumentRecord> DocumentResolver { get; set; }
        public string RecoveryNotice { get; private set; }
        public string SharedVaultPath { get; set; }
        public ProjectStore(string projectDirectory)
        {
            if (string.IsNullOrWhiteSpace(projectDirectory)) throw new ArgumentException("A project directory is required.", nameof(projectDirectory));
            ProjectDirectory = Path.GetFullPath(projectDirectory);
            MetadataPath = Path.Combine(ProjectDirectory, "project.xml");
            DocumentsDirectory = Path.Combine(ProjectDirectory, "documents");
        }
        public string ProjectDirectory { get; }
        public string MetadataPath { get; }
        public string DocumentsDirectory { get; }
        private static XmlSerializer CreateMetadataSerializer()
        {
            var overrides = new XmlAttributeOverrides();
            overrides.Add(typeof(DocumentRecord), nameof(DocumentRecord.IndexedPages), new XmlAttributes { XmlIgnore = true });
            return new XmlSerializer(typeof(ProjectState), overrides);
        }
        public static ProjectState ReadMetadata(Stream input)
        {
            using (var xml = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 256L * 1024 * 1024 }))
            {
                var state = (ProjectState)reader.Deserialize(xml);
                NormalizeState(state); return state;
            }
        }
        public static byte[] MetadataBytes(ProjectState state)
        {
            using (var output = new MemoryStream())
            {
                using (var xml = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), NewLineHandling = NewLineHandling.Entitize })) writer.Serialize(xml, state);
                return output.ToArray();
            }
        }
        public ProjectState LoadOrCreate(string workbookPath)
        {
            lock (sync)
            {
                Directory.CreateDirectory(DocumentsDirectory);
                if (!File.Exists(MetadataPath)) return new ProjectState { WorkbookPath = workbookPath ?? "" };
                ProjectState state;
                try { using (var file = File.OpenRead(MetadataPath)) state = ReadMetadata(file); }
                catch (Exception failure) when (failure is InvalidOperationException || failure is XmlException)
                {
                    if (!File.Exists(MetadataPath + ".bak")) throw;
                    using (var backup = File.OpenRead(MetadataPath + ".bak")) state = ReadMetadata(backup);
                    File.Move(MetadataPath, MetadataPath + ".corrupt-" + Guid.NewGuid().ToString("N"));
                    File.Copy(MetadataPath + ".bak", MetadataPath);
                    RecoveryNotice = "Métadonnées récupérées depuis la sauvegarde précédente. Vérifiez les dernières preuves.";
                }
                state.WorkbookPath = workbookPath ?? state.WorkbookPath;
                SharedVaultPath = state.SharedVaultPath;
                foreach (var document in state.Documents) ConfigureIndex(document);
                return state;
            }
        }
        public void Save(ProjectState state)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            lock (sync)
            {
                Directory.CreateDirectory(DocumentsDirectory);
                foreach (var document in state.Documents)
                {
                    if (document.IndexDirty || (string.IsNullOrEmpty(document.IndexKey) && document.IndexedPages.Count > 0))
                    {
                        var key = Guid.NewGuid().ToString("N");
                        var path = IndexPath(key); Directory.CreateDirectory(Path.GetDirectoryName(path));
                        using (var stream = File.Create(path))
                        using (var zip = new GZipStream(stream, CompressionMode.Compress)) indexSerializer.Serialize(zip, document.IndexedPages);
                        document.IndexKey = key; document.MarkIndexSaved();
                    }
                    ConfigureIndex(document);
                }
                var revision = state.Revision; state.UpdatedAtUtc = DateTime.UtcNow; state.Revision = Guid.NewGuid().ToString("N");
                try { AtomicWrite(MetadataPath, MetadataBytes(state)); }
                catch { state.Revision = revision; throw; }
                SharedVaultPath = state.SharedVaultPath;
                // Append-only metadata checkpoints; originals/indices are immutable and shared.
                var recovery = Path.Combine(ProjectDirectory, "recovery"); Directory.CreateDirectory(recovery);
                try
                {
                    File.Copy(MetadataPath, Path.Combine(recovery, DateTime.UtcNow.ToString("yyyyMMddHHmmssfffffff") + "-" + state.Revision + ".xml"));
                    foreach (var old in Directory.GetFiles(recovery, "*.xml").OrderByDescending(x => x).Skip(20)) File.Delete(old);
                }
                catch (IOException) { /* project.xml and its previous version are already durable */ }
                Saved?.Invoke();
            }
        }
        private void ConfigureIndex(DocumentRecord document)
        {
            if (string.IsNullOrEmpty(document.IndexKey)) return;
            document.PageLoader = () =>
            {
                var path = IndexPath(document.IndexKey);
                if (!File.Exists(path)) { document.IndexComplete = false; return new List<PageTextRecord>(); }
                using (var stream = File.OpenRead(path))
                using (var zip = new GZipStream(stream, CompressionMode.Decompress))
                using (var xml = XmlReader.Create(zip, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 256L * 1024 * 1024 }))
                    return (List<PageTextRecord>)indexSerializer.Deserialize(xml);
            };
            if (!File.Exists(IndexPath(document.IndexKey))) document.IndexComplete = false;
        }
        public string IndexPath(string key)
        {
            if (!Guid.TryParseExact(key, "N", out _)) throw new InvalidDataException("Identifiant d'index incorrect.");
            return Path.Combine(ProjectDirectory, "indexes", key + ".xml.gz");
        }
        public string ResolveDocumentPath(DocumentRecord document)
        {
            var path = LocalDocumentPath(document);
            if (!File.Exists(path)) DocumentResolver?.Invoke(document);
            if (!File.Exists(path) && !string.IsNullOrWhiteSpace(SharedVaultPath))
            {
                var source = SharedVault.DocumentPath(SharedVaultPath, document);
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                var temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try { File.Copy(source, temp); VerifyHash(document, temp); File.Move(temp, path); }
                finally { if (File.Exists(temp)) File.Delete(temp); }
            }
            if(File.Exists(path))
            {
                var info=new FileInfo(path);var stamp=info.Length+"|"+info.LastWriteTimeUtc.Ticks+"|"+document.Sha256;
                if(!verified.TryGetValue(path,out var previous)||previous!=stamp){VerifyHash(document,path);verified[path]=stamp;}
            }
            return path;
        }
        public string LocalDocumentPath(DocumentRecord document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(document.RelativePath) || Path.IsPathRooted(document.RelativePath)) throw new InvalidOperationException("Chemin de pièce invalide.");
            var full = Path.GetFullPath(Path.Combine(ProjectDirectory, document.RelativePath.Replace('\\', Path.DirectorySeparatorChar)));
            if (!full.StartsWith(ProjectDirectory.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Le chemin sort du projet.");
            return full;
        }
        public static void VerifyHash(DocumentRecord document, string path)
        {
            if (!string.Equals(document.Sha256, DocumentImporter.ComputeSha256(path), StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Pièce corrompue ou remplacée : " + document.OriginalName);
        }
        public static void AtomicWrite(string path, byte[] data)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { stream.Write(data, 0, data.Length); stream.Flush(true); }
                if (File.Exists(path)) File.Replace(temp, path, path + ".bak", true); else File.Move(temp, path);
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }
        private static void NormalizeState(ProjectState state)
        {
            if (state == null || state.SchemaVersion > 3) throw new InvalidDataException("Projet invalide ou version plus récente requise.");
            var legacy = state.SchemaVersion < 2; state.SchemaVersion = 3;
            state.CellLinks = state.CellLinks ?? new List<CellLinkRecord>();
            state.Documents = state.Documents ?? new List<DocumentRecord>(); state.Snips = state.Snips ?? new List<SnipRecord>();
            state.AuditTrail = state.AuditTrail ?? new List<AuditEventRecord>(); state.XrefReservations = state.XrefReservations ?? new List<XrefReservation>();
            foreach (var document in state.Documents)
            {
                document.Categories = document.Categories ?? new List<string>();
                if (legacy) document.IndexComplete = document.PageCount > 0 && document.IndexedPages.Count == document.PageCount;
            }
            if (state.Documents.Select(d => d.Id).Distinct().Count() != state.Documents.Count || state.Snips.Select(s => s.Id).Distinct().Count() != state.Snips.Count)
                throw new InvalidDataException("Identifiants de preuves ou pièces en double.");
        }
    }
}
