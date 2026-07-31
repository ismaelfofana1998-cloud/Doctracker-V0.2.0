using System;
using System.Drawing;
using System.IO;
using Tesseract;

namespace Doctracker.AddIn.Infrastructure
{
    internal interface IOcrEngine
    {
        string Recognize(Bitmap bitmap);
    }

    internal sealed class TesseractOcrEngine : IOcrEngine
    {
        private readonly string dataPath;

        public TesseractOcrEngine()
        {
            dataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
        }

        public string Recognize(Bitmap bitmap)
        {
            if (bitmap == null) throw new ArgumentNullException(nameof(bitmap));
            if (!File.Exists(Path.Combine(dataPath, "fra.traineddata")) ||
                !File.Exists(Path.Combine(dataPath, "eng.traineddata")))
            {
                throw new FileNotFoundException(
                    "The French OCR data is missing. Run scripts\\prepare-assets.cmd, then rebuild Doctracker.");
            }

            try
            {
                using (var engine = new TesseractEngine(dataPath, "fra+eng", EngineMode.Default))
                using (var pix = PixConverter.ToPix(bitmap))
                using (var page = engine.Process(pix))
                {
                    return (page.GetText() ?? string.Empty).Trim();
                }
            }
            catch (DllNotFoundException exception)
            {
                throw CreateNativeDependencyException(exception);
            }
            catch (TypeInitializationException exception)
            {
                throw CreateNativeDependencyException(exception);
            }
        }

        private static InvalidOperationException CreateNativeDependencyException(Exception innerException)
        {
            return new InvalidOperationException(
                "Le moteur OCR natif ne peut pas démarrer. Réinstallez Doctracker et vérifiez que les redistribuables Microsoft Visual C++ 2015-2022 x86 et x64 sont installés.",
                innerException);
        }
    }
}
