using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace Doctracker.AddIn.Infrastructure
{
    /// <summary>
    /// Loads PDFium from the deployed add-in directory.
    ///
    /// VSTO add-ins run inside Excel.exe. A plain DllImport("pdfium.dll")
    /// therefore searches from Excel's process directories, not reliably from
    /// the ClickOnce directory that contains Doctracker.AddIn.dll. VSTO also
    /// shadow-copies managed assemblies to AppData\Local\assembly\dl3, so
    /// Assembly.Location does not necessarily identify the deployment folder.
    /// Loading the architecture-specific native library from the assembly's
    /// original CodeBase first makes subsequent PdfiumViewer P/Invoke calls
    /// deterministic.
    /// </summary>
    internal static class NativePdfiumLoader
    {
        private const uint LoadWithAlteredSearchPath = 0x00000008;
        private static readonly object SyncRoot = new object();
        private static IntPtr moduleHandle;

        public static void EnsureLoaded()
        {
            if (moduleHandle != IntPtr.Zero)
            {
                return;
            }

            lock (SyncRoot)
            {
                if (moduleHandle != IntPtr.Zero)
                {
                    return;
                }

                var architecture = Environment.Is64BitProcess ? "x64" : "x86";
                var deploymentDirectories = GetDeploymentDirectories().ToArray();
                var candidates = GetCandidates(deploymentDirectories, architecture).ToArray();
                var pdfiumPath = candidates.FirstOrDefault(File.Exists);

                if (pdfiumPath == null)
                {
                    throw new DllNotFoundException(
                        "Le moteur PDF de Doctracker est absent des dossiers de déploiement " +
                        "détectés. Chemins vérifiés : " +
                        string.Join(" ; ", candidates.Take(8)) +
                        ". Réinstallez Doctracker avec l'artefact Windows Installer complet.");
                }

                moduleHandle = LoadLibraryEx(
                    pdfiumPath,
                    IntPtr.Zero,
                    LoadWithAlteredSearchPath);

                if (moduleHandle == IntPtr.Zero)
                {
                    var errorCode = Marshal.GetLastWin32Error();
                    throw new DllNotFoundException(
                        "Le moteur PDF de Doctracker existe mais Windows ne peut pas le charger. " +
                        "Fichier : " + pdfiumPath + ". Code Windows : " + errorCode + ". " +
                        GetLoadFailureHint(errorCode, architecture));
                }
            }
        }

        private static IEnumerable<string> GetDeploymentDirectories()
        {
            var assembly = typeof(NativePdfiumLoader).Assembly;
            var directories = new List<string>();

            // CodeBase keeps the assembly's original deployment URI when VSTO
            // loads a shadow copy from AppData\Local\assembly\dl3.
            AddCodeBaseDirectory(directories, assembly.CodeBase);
            AddFileDirectory(directories, assembly.Location);
            AddDirectory(directories, AppDomain.CurrentDomain.SetupInformation.ApplicationBase);
            AddDirectory(directories, AppDomain.CurrentDomain.BaseDirectory);

            if (directories.Count == 0)
            {
                throw new DllNotFoundException(
                    "Doctracker ne peut pas déterminer son dossier de déploiement.");
            }

            return directories;
        }

        private static void AddCodeBaseDirectory(
            ICollection<string> directories,
            string codeBase)
        {
            if (string.IsNullOrWhiteSpace(codeBase))
            {
                return;
            }

            try
            {
                var uri = new Uri(codeBase, UriKind.Absolute);
                if (uri.IsFile)
                {
                    AddFileDirectory(directories, uri.LocalPath);
                }
            }
            catch (UriFormatException)
            {
                AddFileDirectory(directories, codeBase);
            }
        }

        private static void AddFileDirectory(
            ICollection<string> directories,
            string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            try
            {
                AddDirectory(directories, Path.GetDirectoryName(filePath));
            }
            catch (ArgumentException)
            {
                // Ignore malformed fallback paths and continue with the other
                // deployment location sources.
            }
            catch (NotSupportedException)
            {
                // Ignore malformed fallback paths and continue with the other
                // deployment location sources.
            }
        }

        private static void AddDirectory(
            ICollection<string> directories,
            string directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                return;
            }

            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(directory)
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch (ArgumentException)
            {
                return;
            }
            catch (NotSupportedException)
            {
                return;
            }
            catch (IOException)
            {
                return;
            }

            if (!directories.Any(
                    candidate => string.Equals(
                        candidate,
                        fullPath,
                        StringComparison.OrdinalIgnoreCase)))
            {
                directories.Add(fullPath);
            }
        }

        private static IEnumerable<string> GetCandidates(
            IEnumerable<string> deploymentDirectories,
            string architecture)
        {
            var emitted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var deploymentDirectory in deploymentDirectories)
            {
                var directCandidates = new[]
                {
                    Path.Combine(deploymentDirectory, architecture, "pdfium.dll"),
                    Path.Combine(deploymentDirectory, "pdfium.dll")
                };

                foreach (var candidate in directCandidates)
                {
                    if (emitted.Add(candidate))
                    {
                        yield return candidate;
                    }
                }

                // Defensive fallback for ClickOnce layouts that retain an
                // additional package directory while preserving x86/x64.
                foreach (var discovered in DiscoverArchitectureSpecificLibraries(
                             deploymentDirectory,
                             architecture))
                {
                    if (emitted.Add(discovered))
                    {
                        yield return discovered;
                    }
                }
            }
        }

        private static IEnumerable<string> DiscoverArchitectureSpecificLibraries(
            string deploymentDirectory,
            string architecture)
        {
            if (!Directory.Exists(deploymentDirectory))
            {
                return Enumerable.Empty<string>();
            }

            try
            {
                return Directory
                    .EnumerateFiles(
                        deploymentDirectory,
                        "pdfium.dll",
                        SearchOption.AllDirectories)
                    .Where(path =>
                        string.Equals(
                            new DirectoryInfo(Path.GetDirectoryName(path)).Name,
                            architecture,
                            StringComparison.OrdinalIgnoreCase))
                    .ToArray();
            }
            catch (UnauthorizedAccessException)
            {
                return Enumerable.Empty<string>();
            }
            catch (IOException)
            {
                return Enumerable.Empty<string>();
            }
        }

        private static string GetLoadFailureHint(int errorCode, string architecture)
        {
            switch (errorCode)
            {
                case 126:
                    return "Une dépendance native est manquante. Installez le redistribuable " +
                           "Microsoft Visual C++ 2015-2022 " + architecture + ", puis relancez Excel.";
                case 193:
                    return "L'architecture ne correspond pas à Excel. Doctracker a sélectionné " +
                           architecture + " d'après le processus Excel en cours.";
                case 5:
                    return "Windows refuse l'accès au fichier. Vérifiez que l'artefact a été " +
                           "entièrement décompressé et débloqué avant l'installation.";
                default:
                    return "Fermez Excel, réinstallez Doctracker, puis réessayez.";
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadLibraryEx(
            string fileName,
            IntPtr file,
            uint flags);
    }
}
