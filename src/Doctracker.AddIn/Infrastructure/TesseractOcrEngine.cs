using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Doctracker.Core.Models;
using Tesseract;

namespace Doctracker.AddIn.Infrastructure
{
    internal interface IOcrEngine : IDisposable
    {
        PageTextRecord Recognize(Bitmap bitmap, bool table = false);
    }

    internal sealed class TesseractOcrEngine : IOcrEngine
    {
        private readonly object sync = new object();
        private TesseractEngine engine;
        private bool disposed;

        public PageTextRecord Recognize(Bitmap bitmap, bool table = false)
        {
            if (bitmap == null) throw new ArgumentNullException(nameof(bitmap));
            lock (sync)
            {
                if (disposed) throw new ObjectDisposedException(nameof(TesseractOcrEngine));
                try
                {
                    EnsureEngine();
                    using (var pix = PixConverter.ToPix(bitmap))
                    using (var page = engine.Process(pix, table ? PageSegMode.Auto : PageSegMode.Auto))
                    using (var iterator = page.GetIterator())
                    {
                        var result = new PageTextRecord { Text = (page.GetText() ?? "").Trim() };
                        var line = 0;
                        iterator.Begin();
                        do
                        {
                            if (iterator.IsAtBeginningOf(PageIteratorLevel.TextLine)) line++;
                            var text = iterator.GetText(PageIteratorLevel.Word);
                            Rect bounds;
                            if (string.IsNullOrWhiteSpace(text) || !iterator.TryGetBoundingBox(PageIteratorLevel.Word, out bounds)) continue;
                            var left = Math.Max(0, bounds.X1) / (double)bitmap.Width;
                            var top = Math.Max(0, bounds.Y1) / (double)bitmap.Height;
                            var right = Math.Min(bitmap.Width, bounds.X2) / (double)bitmap.Width;
                            var bottom = Math.Min(bitmap.Height, bounds.Y2) / (double)bitmap.Height;
                            if (right <= left || bottom <= top) continue;
                            result.Words.Add(new WordRecord { Text = text.Trim(), Line = line,
                                X = left, Y = top, Width = right - left, Height = bottom - top });
                        } while (iterator.Next(PageIteratorLevel.Word));
                        return result;
                    }
                }
                catch (DllNotFoundException exception) { throw NativeFailure(exception); }
                catch (TypeInitializationException exception) { throw NativeFailure(exception); }
            }
        }

        private void EnsureEngine()
        {
            if (engine != null) return;
            var root = NativePdfiumLoader.GetDeploymentDirectories().FirstOrDefault(directory =>
                File.Exists(Path.Combine(directory, "tessdata", "fra.traineddata")) &&
                File.Exists(Path.Combine(directory, "tessdata", "eng.traineddata")));
            if (root == null) throw new FileNotFoundException("Modèles OCR français/anglais introuvables. Réinstallez l'installateur complet.");
            // Tesseract appends x86/x64 itself; point it at the actual ClickOnce deployment.
            TesseractEnviornment.CustomSearchPath = root;
            engine = new TesseractEngine(Path.Combine(root, "tessdata"), "fra+eng", EngineMode.Default);
            engine.SetVariable("preserve_interword_spaces", "1");
        }

        private static Exception NativeFailure(Exception inner) => new InvalidOperationException(
            "Le moteur OCR ne peut pas démarrer. Vérifiez les redistribuables Microsoft Visual C++ 2015-2022 x86 et x64.", inner);

        public void Dispose()
        {
            lock (sync) { disposed = true; engine?.Dispose(); engine = null; }
        }
    }
}
