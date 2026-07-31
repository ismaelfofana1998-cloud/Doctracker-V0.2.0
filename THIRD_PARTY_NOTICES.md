# Third-party components

Doctracker V0.2.0 references the following packages at build time:

- `Tesseract` 5.2.0 — OCR wrapper and native engine dependencies;
- `tessdata_fast` French and English language data;
- `PdfiumViewer` 2.13.0 — PDF rendering;
- `PdfiumViewer.Native.x86.v8-xfa` and
  `PdfiumViewer.Native.x86_64.v8-xfa` 2018.4.8.256 — native PDFium binaries;
- xUnit, Microsoft.NET.Test.Sdk and coverlet — automated tests only.

Their own licenses and notices remain applicable. Before commercial
distribution, the release process must copy the exact license texts shipped by
the restored package versions into the installer.
