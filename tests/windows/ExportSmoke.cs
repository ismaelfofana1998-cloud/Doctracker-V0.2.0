using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using Doctracker.Core.Models;

// Runs with the same managed dependency redirects shipped with the VSTO add-in.
internal static class ExportSmoke
{
    private static int Main(string[] args)
    {
        try
        {
            var document = new DocumentRecord { OriginalName = "native.pdf", TestReference = "DAC B 30 040", ReferenceNumber = 1 };
            document.Comments.Add(new DocumentComment { PageNumber=1, X=.12, Y=.55, Width=.55, Height=.22, Text="Rapprochement validé — document reçu." });
            var snip = new SnipRecord {
                PageNumber = 1, WorksheetName = "Achats", CellAddress = "B2",
                X = double.Parse(args[2], CultureInfo.InvariantCulture),
                Y = double.Parse(args[3], CultureInfo.InvariantCulture),
                Width = double.Parse(args[4], CultureInfo.InvariantCulture),
                Height = double.Parse(args[5], CultureInfo.InvariantCulture)
            };
            var assembly = Assembly.LoadFrom(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Doctracker.AddIn.dll"));
            var exporter = assembly.GetType("Doctracker.AddIn.Infrastructure.AnnotatedPdfExporter", true);
            exporter.GetMethod("Export", BindingFlags.Public | BindingFlags.Static).Invoke(null,
                new object[] { args[0], document, new[] { snip }, args[1], CancellationToken.None });
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }
}
