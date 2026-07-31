# OCR language data

The release workflow downloads `fra.traineddata` and `eng.traineddata` from the
official, signed `tesseract-ocr/tessdata_fast` 4.1.0 release before compilation.
The script retries transient downloads and rejects incomplete files.

For a local Visual Studio build, run `scripts\prepare-assets.cmd` once.
