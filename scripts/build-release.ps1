$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = Split-Path -Parent $PSScriptRoot
$solutionPath = Join-Path $projectRoot "Doctracker.sln"
$addInProject = Join-Path $projectRoot "src\Doctracker.AddIn\Doctracker.AddIn.csproj"
$testProject = Join-Path $projectRoot "tests\Doctracker.Core.Tests\Doctracker.Core.Tests.csproj"
$coreAssembly = Join-Path $projectRoot "src\Doctracker.Core\bin\Release\net48\Doctracker.Core.dll"
$publishDirectory = Join-Path $projectRoot "artifacts\installer"
$env:MSBUILDDISABLENODEREUSE = "1"

function Find-MSBuild {
    $candidate = Get-Command msbuild.exe -ErrorAction SilentlyContinue
    if ($candidate) {
        return $candidate.Source
    }

    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere)) {
        throw "MSBuild est introuvable. Installez Visual Studio 2022 avec la charge Développement Office/SharePoint."
    }

    $installationPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
    if (-not $installationPath) {
        throw "Aucune installation Visual Studio compatible n'a été trouvée."
    }

    $resolved = Join-Path $installationPath "MSBuild\Current\Bin\MSBuild.exe"
    if (-not (Test-Path $resolved)) {
        throw "MSBuild est absent de l'installation Visual Studio détectée."
    }
    return $resolved
}

function Get-DoctrackerCertificate {
    $certificate = Get-ChildItem Cert:\CurrentUser\My |
        Where-Object {
            $_.Subject -eq "CN=Doctracker Development" -and
            $_.HasPrivateKey -and
            $_.NotAfter -gt (Get-Date).AddDays(1)
        } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1

    if (-not $certificate) {
        $certificate = New-SelfSignedCertificate `
            -Type CodeSigningCert `
            -Subject "CN=Doctracker Development" `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -KeyAlgorithm RSA `
            -KeyLength 2048 `
            -HashAlgorithm SHA256 `
            -KeyExportPolicy NonExportable `
            -NotAfter (Get-Date).AddDays(30)
    }
    return $certificate
}

function Invoke-CheckedProcess {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$FilePath s'est terminé avec le code $LASTEXITCODE."
    }
}

function Assert-InstallerContents {
    param([Parameter(Mandatory = $true)][string]$Root)

    $checks = @(
        @{ Label = "setup.exe"; Pattern = "setup.exe"; Minimum = 1 },
        @{ Label = "certificat public"; Pattern = "Doctracker-Development.cer"; Minimum = 1 },
        @{ Label = "assistant d'installation CMD"; Pattern = "Install-Doctracker.cmd"; Minimum = 1 },
        @{ Label = "assistant d'installation PowerShell"; Pattern = "Install-Doctracker.ps1"; Minimum = 1 },
        @{ Label = "manifest .vsto"; Pattern = "*.vsto"; Minimum = 1 },
        @{ Label = "manifest applicatif"; Pattern = "*.dll.manifest*"; Minimum = 1 },
        @{ Label = "Doctracker.AddIn.dll"; Pattern = "Doctracker.AddIn.dll*"; Minimum = 1 },
        @{ Label = "fra.traineddata"; Pattern = "fra.traineddata*"; Minimum = 1 },
        @{ Label = "eng.traineddata"; Pattern = "eng.traineddata*"; Minimum = 1 },
        @{ Label = "pdfium.dll x86/x64"; Pattern = "pdfium.dll*"; Minimum = 2 },
        @{ Label = "tesseract50.dll x86/x64"; Pattern = "tesseract50.dll*"; Minimum = 2 },
        @{ Label = "leptonica native DLL x86/x64"; Pattern = "*leptonica*.dll*"; Minimum = 2 }
    )

    $missing = foreach ($check in $checks) {
        $count = @(Get-ChildItem $Root -Recurse -File -Filter $check.Pattern).Count
        if ($count -lt $check.Minimum) {
            $check.Label
        }
    }
    if ($missing) {
        throw "Installateur incomplet. Fichier(s) absent(s) : $($missing -join ', ')"
    }

    $pdfiumFiles = @(Get-ChildItem $Root -Recurse -File -Filter "pdfium.dll*")
    $missingPdfiumArchitectures = @(
        if (-not ($pdfiumFiles | Where-Object { $_.FullName -match '[\\/]x86[\\/]' })) { "x86" }
        if (-not ($pdfiumFiles | Where-Object { $_.FullName -match '[\\/]x64[\\/]' })) { "x64" }
    )
    if ($missingPdfiumArchitectures.Count -gt 0) {
        throw "Installateur incomplet. PDFium absent pour : $($missingPdfiumArchitectures -join ', ')"
    }

    $applicationManifest = Get-ChildItem $Root -Recurse -File -Filter "*.dll.manifest*" |
        Select-Object -First 1
    $manifestText = [System.IO.File]::ReadAllText($applicationManifest.FullName).Replace("/", "\")
    foreach ($architecture in @("x86", "x64")) {
        if ($manifestText -notmatch "$architecture\\pdfium\.dll(?:\.deploy)?") {
            throw "Le manifeste ClickOnce ne référence pas $architecture\pdfium.dll."
        }
    }

    $setup = (Get-ChildItem $Root -Recurse -File -Filter "setup.exe" | Select-Object -First 1).FullName
    if (-not (Get-AuthenticodeSignature $setup).SignerCertificate) {
        throw "setup.exe n'est pas signé."
    }
}

$msbuild = Find-MSBuild
$certificate = Get-DoctrackerCertificate
$signingProperties = @(
    "/p:Configuration=Release",
    "/p:SignManifests=true",
    "/p:ManifestCertificateThumbprint=$($certificate.Thumbprint)",
    "/p:CertificateThumbprint=$($certificate.Thumbprint)"
)

Invoke-CheckedProcess -FilePath (Join-Path $PSScriptRoot "prepare-assets.cmd") -Arguments @()
Invoke-CheckedProcess -FilePath $msbuild -Arguments @($solutionPath, "/t:Restore", "/p:RestorePackagesConfig=true")
Invoke-CheckedProcess -FilePath "dotnet.exe" -Arguments @(
    "test", $testProject, "--framework", "net48", "--configuration", "Release", "--no-restore"
)
if (-not (Test-Path $coreAssembly)) {
    throw "Doctracker.Core.dll net48 est absent. Le moteur doit cibler .NET Framework 4.8 pour rester compatible avec la tâche VSTO FindRibbons."
}
Invoke-CheckedProcess -FilePath $msbuild -Arguments (
    @($addInProject, "/m:1", "/nr:false", "/t:Build") + $signingProperties
)

if (Test-Path $publishDirectory) {
    Remove-Item $publishDirectory -Recurse -Force
}

$publishArguments = @(
    $addInProject, "/m:1", "/nr:false", "/t:Publish", "/p:PublishDir=$publishDirectory\"
) + $signingProperties
Invoke-CheckedProcess -FilePath $msbuild -Arguments $publishArguments

$publicCertificatePath = Join-Path $publishDirectory "Doctracker-Development.cer"
Export-Certificate -Cert $certificate -FilePath $publicCertificatePath -Type CERT | Out-Null
Copy-Item `
    -Path (Join-Path $projectRoot "installer\*") `
    -Destination $publishDirectory `
    -Force

$exportedCertificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($publicCertificatePath)
if ($exportedCertificate.HasPrivateKey) {
    throw "Le certificat public exporté contient une clé privée."
}
$setupPath = (Get-ChildItem $publishDirectory -Recurse -File -Filter "setup.exe" | Select-Object -First 1).FullName
$setupSignature = Get-AuthenticodeSignature $setupPath
if ($null -eq $setupSignature.SignerCertificate -or
    $setupSignature.SignerCertificate.Thumbprint -ne $exportedCertificate.Thumbprint) {
    throw "Le certificat public ne correspond pas à la signature de setup.exe."
}

Assert-InstallerContents -Root $publishDirectory

Write-Host "Build, tests et installateur terminés : $publishDirectory"
