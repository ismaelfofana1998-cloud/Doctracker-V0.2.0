using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public sealed class DocumentImporter
    {
        private static readonly string[] AllowedExtensions =
        {
            ".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"
        };

        private readonly ProjectStore store;

        public DocumentImporter(ProjectStore store)
        {
            this.store = store ?? throw new ArgumentNullException(nameof(store));
        }

        public DocumentRecord Import(ProjectState state, string sourcePath, string actor, string category = null, bool persist = true, string sourceName = null, string sourceHash = null)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("Document not found.", sourcePath);

            var extension = Path.GetExtension(sourcePath).ToLowerInvariant();
            if (!AllowedExtensions.Contains(extension))
            {
                throw new NotSupportedException("Supported formats: PDF, PNG, JPG, TIFF and BMP.");
            }

            var hash = ComputeSha256(sourcePath);
            var duplicate = state.Documents.FirstOrDefault(d => d.Sha256 == hash);
            if (duplicate != null)
            {
                // Re-import can repair an accidentally removed local copy.
                var existingPath = store.ResolveDocumentPath(duplicate);
                if (!File.Exists(existingPath))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(existingPath));
                    File.Copy(sourcePath, existingPath, false);
                }
                var previousImport = duplicate.LastImportedAtUtc;
                var previousCategories = duplicate.Categories.ToList();
                duplicate.LastImportedAtUtc = DateTime.UtcNow;
                AddCategory(duplicate, category);
                try { if (persist) store.Save(state); }
                catch { duplicate.LastImportedAtUtc = previousImport; duplicate.Categories = previousCategories; throw; }
                return duplicate;
            }

            Directory.CreateDirectory(store.DocumentsDirectory);
            var id = Guid.NewGuid().ToString("N");
            var safeName = id + extension;
            var destination = Path.Combine(store.DocumentsDirectory, safeName);
            File.Copy(sourcePath, destination, false);

            var document = new DocumentRecord
            {
                Id = id,
                OriginalName = Path.GetFileName(sourcePath),
                RelativePath = Path.Combine("documents", safeName),
                Sha256 = hash,
                ByteLength = new FileInfo(destination).Length,
                LastImportedAtUtc = DateTime.UtcNow,
                OriginalSourceName = sourceName ?? string.Empty, OriginalSourceHash = sourceHash ?? string.Empty,
                AddedAtUtc = DateTime.UtcNow
            };

            AddCategory(document, category);
            state.Documents.Add(document);
            state.AuditTrail.Add(new AuditEventRecord
            {
                Actor = actor ?? string.Empty,
                Action = "DocumentImported",
                EntityType = "Document",
                EntityId = document.Id,
                Details = document.OriginalName
            });
            try { if (persist) store.Save(state); }
            catch
            {
                state.Documents.Remove(document);
                state.AuditTrail.RemoveAt(state.AuditTrail.Count - 1);
                File.Delete(destination);
                throw;
            }
            return document;
        }

        public static bool IsSupported(string path) => AllowedExtensions.Contains(Path.GetExtension(path).ToLowerInvariant());
        public static void AddCategory(DocumentRecord document, string category)
        {
            if (!string.IsNullOrWhiteSpace(category) && !document.Categories.Contains(category.Trim(), StringComparer.OrdinalIgnoreCase)) document.Categories.Add(category.Trim());
        }
        public static string ComputeSha256(string path)
        {
            using (var stream = File.OpenRead(path))
            using (var sha = SHA256.Create())
            {
                return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", string.Empty);
            }
        }
    }
}
