using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using Doctracker.Core.Models;

namespace Doctracker.Core.Services
{
    public static class SharedVault
    {
        public static string DocumentPath(string root, DocumentRecord document)
        {
            if (document.Sha256 == null || document.Sha256.Length != 64 || document.Sha256.Any(c => !Uri.IsHexDigit(c))) throw new InvalidDataException("Empreinte de pièce incorrecte.");
            return Path.Combine(root, "objects", document.Sha256.ToUpperInvariant() + ".bin");
        }
        public static void Publish(string root, ProjectStore store, ProjectState state)
        {
            Directory.CreateDirectory(Path.Combine(root, "objects"));
            foreach (var document in state.Documents)
            {
                var destination = DocumentPath(root, document);
                if (File.Exists(destination)) { if(document.ByteLength>0 && new FileInfo(destination).Length!=document.ByteLength)throw new InvalidDataException("Pièce partagée altérée : "+document.OriginalName); continue; }
                var source = store.ResolveDocumentPath(document); ProjectStore.VerifyHash(document, source);
                var temp = destination + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    File.Copy(source, temp);
                    try { File.Move(temp, destination); }
                    catch (IOException) { if (!File.Exists(destination)) throw; ProjectStore.VerifyHash(document, destination); }
                }
                finally { if (File.Exists(temp)) File.Delete(temp); }
            }
        }
        private static string ReservationFolder(string root,string projectId,string reference)
        {
            using(var hash=SHA256.Create())return Path.Combine(root,"xref",BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(projectId+"|"+reference.Trim().ToUpperInvariant()))).Replace("-",""));
        }
        public static System.Collections.Generic.IEnumerable<int> ReservedNumbers(string root,string projectId,string reference)
        {
            var folder=ReservationFolder(root,projectId,reference);
            return !Directory.Exists(folder)?new int[0]:Directory.GetFiles(folder,"*.txt").Select(path=>int.TryParse(Path.GetFileNameWithoutExtension(path),out var number)?number:0).Where(number=>number>0).ToArray();
        }
        public static void Reserve(string root, string projectId, string reference, int number, string documentId)
        {
            var folder=ReservationFolder(root,projectId,reference); Directory.CreateDirectory(folder);
            var path = Path.Combine(folder, number.ToString("D6") + ".txt");
            try { using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { var data = Encoding.UTF8.GetBytes(documentId); file.Write(data,0,data.Length); file.Flush(true); } }
            catch (IOException) { if (!File.Exists(path) || File.ReadAllText(path) != documentId) throw new IOException("Ce numéro Xref est déjà réservé. Choisissez le suivant."); }
        }
    }
    public static class CrossReferences
    {
        public static int Next(ProjectState state, string reference) => Available(state,reference).First();
        public static System.Collections.Generic.IEnumerable<int> Available(ProjectState state,string reference)
        {
            var used=new System.Collections.Generic.HashSet<int>(state.XrefReservations.Where(r=>string.Equals(r.Reference,reference.Trim(),StringComparison.OrdinalIgnoreCase)).Select(r=>r.Number));
            if(!string.IsNullOrEmpty(state.SharedVaultPath))used.UnionWith(SharedVault.ReservedNumbers(state.SharedVaultPath,state.ProjectId,reference));
            return Enumerable.Range(1,99999).Where(n=>!used.Contains(n));
        }
        public static bool IsUsed(ProjectState state, string reference, int number) => state.XrefReservations.Any(r => r.Number == number && string.Equals(r.Reference,reference.Trim(),StringComparison.OrdinalIgnoreCase));
        public static void Assign(ProjectState state, DocumentRecord document, string reference, int number)
        {
            reference = (reference ?? "").Trim();
            if (reference.Length == 0 || reference.Length > 80 || number < 1 || number > 99999) throw new ArgumentException("Renseignez une référence de test (80 caractères maximum) et un numéro entre 1 et 99999.");
            if (IsUsed(state, reference, number)) throw new InvalidOperationException("Ce numéro Xref est déjà utilisé.");
            if (!string.IsNullOrEmpty(state.SharedVaultPath)) SharedVault.Reserve(state.SharedVaultPath,state.ProjectId,reference,number,document.Id);
            state.XrefReservations.Add(new XrefReservation { Reference=reference, Number=number, DocumentId=document.Id });
            document.TestReference=reference; document.ReferenceNumber=number;
        }
        public static string SafeFileName(string value)
        {
            var name = new string((value ?? "document").Select(c => c < 32 || "<>:\"/\\|?*".Contains(c) ? '_' : c).ToArray()).Trim().TrimEnd('.');
            return name.Length > 180 ? name.Substring(0,180) : name.Length == 0 ? "document" : name;
        }
    }
}
