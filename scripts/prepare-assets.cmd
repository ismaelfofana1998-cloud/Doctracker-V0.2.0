@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PROJECT_DIR=%SCRIPT_DIR%.."
set "TESSDATA_DIR=%PROJECT_DIR%\src\Doctracker.AddIn\tessdata"

if not exist "%TESSDATA_DIR%" mkdir "%TESSDATA_DIR%"

echo Downloading French OCR data...
curl.exe -fL --retry 3 --retry-all-errors "https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/4.1.0/fra.traineddata" -o "%TESSDATA_DIR%\fra.traineddata"
if errorlevel 1 exit /b 1

echo Downloading English OCR data...
curl.exe -fL --retry 3 --retry-all-errors "https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/4.1.0/eng.traineddata" -o "%TESSDATA_DIR%\eng.traineddata"
if errorlevel 1 exit /b 1

for %%F in ("%TESSDATA_DIR%\fra.traineddata" "%TESSDATA_DIR%\eng.traineddata") do (
  if %%~zF LSS 100000 (
    echo OCR asset %%~nxF is unexpectedly small.
    exit /b 1
  )
)

echo OCR assets are ready.
endlocal
