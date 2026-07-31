using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public sealed class ProjectStore
    {
        private readonly object sync = new object();
        private readonly XmlSerializer serializer = new XmlSerializer(typeof(ProjectState));

        public ProjectStore(string projectDirectory)
        {
            if (string.IsNullOrWhiteSpace(projectDirectory))
            {
                throw new ArgumentException("A project directory is required.", nameof(projectDirectory));
            }

            ProjectDirectory = Path.GetFullPath(projectDirectory);
            MetadataPath = Path.Combine(ProjectDirectory, "project.xml");
            DocumentsDirectory = Path.Combine(ProjectDirectory, "documents");
        }

        public string ProjectDirectory { get; }
        public string MetadataPath { get; }
        public string DocumentsDirectory { get; }

        public ProjectState LoadOrCreate(string workbookPath)
        {
            lock (sync)
            {
                Directory.CreateDirectory(ProjectDirectory);
                Directory.CreateDirectory(DocumentsDirectory);

                if (!File.Exists(MetadataPath))
                {
                    return new ProjectState { WorkbookPath = workbookPath ?? string.Empty };
                }

                using (var stream = File.OpenRead(MetadataPath))
                {
                    var state = (ProjectState)serializer.Deserialize(stream);
                    NormalizeState(state);
                    state.WorkbookPath = workbookPath ?? state.WorkbookPath;
                    return state;
                }
            }
        }

        public void Save(ProjectState state)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }

            lock (sync)
            {
                Directory.CreateDirectory(ProjectDirectory);
                Directory.CreateDirectory(DocumentsDirectory);
                state.UpdatedAtUtc = DateTime.UtcNow;

                var tempPath = MetadataPath + ".tmp";
                var settings = new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false),
                    Indent = true,
                    NewLineHandling = NewLineHandling.Entitize
                };

                using (var writer = XmlWriter.Create(tempPath, settings))
                {
                    serializer.Serialize(writer, state);
                }

                if (File.Exists(MetadataPath))
                {
                    File.Replace(tempPath, MetadataPath, MetadataPath + ".bak", true);
                }
                else
                {
                    File.Move(tempPath, MetadataPath);
                }
            }
        }

        public string ResolveDocumentPath(DocumentRecord document)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            var fullPath = Path.GetFullPath(Path.Combine(ProjectDirectory, document.RelativePath));
            var root = ProjectDirectory.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("The stored document path escapes the project directory.");
            }

            return fullPath;
        }

        private static void NormalizeState(ProjectState state)
        {
            if (state == null)
            {
                throw new InvalidDataException("The Doctracker project file is empty or invalid.");
            }

            state.ProjectId = state.ProjectId ?? Guid.NewGuid().ToString("N");
            state.WorkbookPath = state.WorkbookPath ?? string.Empty;
            state.Documents = state.Documents ?? new System.Collections.Generic.List<DocumentRecord>();
            state.Snips = state.Snips ?? new System.Collections.Generic.List<SnipRecord>();
            state.AuditTrail = state.AuditTrail ?? new System.Collections.Generic.List<AuditEventRecord>();

            foreach (var document in state.Documents)
            {
                document.Id = document.Id ?? Guid.NewGuid().ToString("N");
                document.OriginalName = document.OriginalName ?? string.Empty;
                document.RelativePath = document.RelativePath ?? string.Empty;
                document.Sha256 = document.Sha256 ?? string.Empty;
                document.IndexedPages = document.IndexedPages ??
                    new System.Collections.Generic.List<PageTextRecord>();
            }

            foreach (var snip in state.Snips)
            {
                snip.Id = snip.Id ?? Guid.NewGuid().ToString("N");
                snip.DocumentId = snip.DocumentId ?? string.Empty;
                snip.RawText = snip.RawText ?? string.Empty;
                snip.ExtractedValue = snip.ExtractedValue ?? string.Empty;
                snip.WorksheetName = snip.WorksheetName ?? string.Empty;
                snip.CellAddress = snip.CellAddress ?? string.Empty;
                snip.Comment = snip.Comment ?? string.Empty;
                snip.PreparedBy = snip.PreparedBy ?? string.Empty;
                snip.ReviewedBy = snip.ReviewedBy ?? string.Empty;
            }
        }
    }
}
