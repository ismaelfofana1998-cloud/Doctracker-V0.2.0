using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Doctracker.Core.Models;
using Doctracker.Core.Services;

namespace Doctracker.AddIn.Infrastructure
{
    internal static class WordDocumentImporter
    {
        public static bool IsWord(string path) => new[] { ".doc", ".docx" }.Contains(Path.GetExtension(path).ToLowerInvariant());
        public static bool IsSupported(string path) => IsWord(path) || DocumentImporter.IsSupported(path);
        public static DocumentRecord Import(WorkbookProjectContext context, string path, string category, bool persist, CancellationToken cancellation)
        {
            cancellation.ThrowIfCancellationRequested();
            if (!IsWord(path)) return context.Importer.Import(context.State, path, Environment.UserName, category, persist);
            var hash = DocumentImporter.ComputeSha256(path);
            var existing = context.State.Documents.FirstOrDefault(d => d.OriginalSourceHash == hash);
            if (existing != null)
                return context.Importer.Import(context.State, context.Store.ResolveDocumentPath(existing), Environment.UserName, category, persist);
            var folder = Path.Combine(Path.GetTempPath(), "Doctracker-Word-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            try
            {
                var output = Path.Combine(folder, Path.GetFileNameWithoutExtension(path) + ".pdf");
                ConvertToPdf(path, output, cancellation);
                cancellation.ThrowIfCancellationRequested();
                return context.Importer.Import(context.State, output, Environment.UserName, category, persist, Path.GetFileName(path), hash);
            }
            finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
        }
        public static void ConvertToPdf(string source, string destination, CancellationToken cancellation)
        {
            Exception failure = null;
            var thread = new Thread(() =>
            {
                dynamic word = null, documents = null, document = null, options = null;
                try
                {
                    cancellation.ThrowIfCancellationRequested();
                    var type = Type.GetTypeFromProgID("Word.Application");
                    if (type == null) throw new InvalidOperationException("L'import .doc/.docx nécessite Microsoft Word installé. Vous pouvez aussi exporter le fichier en PDF puis l'importer.");
                    word = Activator.CreateInstance(type);
                    word.Visible = false; word.DisplayAlerts = 0; word.AutomationSecurity = 3;
                    options = word.Options; options.UpdateLinksAtOpen = false; options.SaveNormalPrompt = false;
                    documents = word.Documents;
                    document = documents.Open(FileName: Path.GetFullPath(source), ConfirmConversions: false, ReadOnly: true,
                        AddToRecentFiles: false, Visible: false, NoEncodingDialog: true,
                        PasswordDocument: Guid.NewGuid().ToString("N"), PasswordTemplate: Guid.NewGuid().ToString("N"));
                    cancellation.ThrowIfCancellationRequested();
                    document.ExportAsFixedFormat(OutputFileName: destination, ExportFormat: 17, OpenAfterExport: false, KeepIRM: true);
                    cancellation.ThrowIfCancellationRequested();
                }
                catch (Exception exception) { failure = exception; }
                finally
                {
                    try { if (document != null) document.Close(SaveChanges: 0); } catch { }
                    try { if (word != null) word.Quit(SaveChanges: 0); } catch { }
                    foreach (object item in new object[] { document, documents, options, word })
                        if (item != null && Marshal.IsComObject(item)) try { Marshal.FinalReleaseComObject(item); } catch (InvalidComObjectException) { }
                }
            }) { IsBackground = true, Name = "Doctracker Word conversion" };
            thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();
            if (failure is OperationCanceledException) throw failure;
            if (failure != null) throw new InvalidOperationException("Conversion Word impossible : " + failure.Message, failure);
            if (!File.Exists(destination)) throw new IOException("Word n'a pas créé le PDF attendu.");
        }
    }
}
