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

        public DocumentRecord Import(ProjectState state, string sourcePath, string actor)
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
                AddedAtUtc = DateTime.UtcNow
            };

            state.Documents.Add(document);
            state.AuditTrail.Add(new AuditEventRecord
            {
                Actor = actor ?? string.Empty,
                Action = "DocumentImported",
                EntityType = "Document",
                EntityId = document.Id,
                Details = document.OriginalName
            });
            store.Save(state);
            return document;
        }

        private static string ComputeSha256(string path)
        {
            using (var stream = File.OpenRead(path))
            using (var sha = SHA256.Create())
            {
                return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", string.Empty);
            }
        }
    }
}
