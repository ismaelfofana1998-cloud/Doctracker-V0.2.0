$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$packageDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$certificatePath = Join-Path $packageDirectory "Doctracker-Development.cer"
$setupPath = Join-Path $packageDirectory "setup.exe"
$expectedSubject = "CN=Doctracker Development"
$codeSigningOid = "1.3.6.1.5.5.7.3.3"

function Close-CertificateStore {
    param([System.Security.Cryptography.X509Certificates.X509Store]$Store)

    if ($null -ne $Store) {
        $Store.Close()
    }
}

if (-not (Test-Path -LiteralPath $certificatePath -PathType Leaf)) {
    throw "Le certificat public Doctracker-Development.cer est absent du dossier d'installation."
}
if (-not (Test-Path -LiteralPath $setupPath -PathType Leaf)) {
    throw "setup.exe est absent du dossier d'installation."
}

$certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certificatePath)
$now = Get-Date
if ($certificate.Subject -ne $expectedSubject) {
    throw "Sujet de certificat inattendu : $($certificate.Subject)."
}
if ($certificate.HasPrivateKey) {
    throw "Le package ne doit jamais contenir la clé privée du certificat."
}
if ($certificate.NotBefore -gt $now -or $certificate.NotAfter -le $now) {
    throw "Le certificat de cette compilation n'est pas actuellement valide."
}

$enhancedKeyUsage = @($certificate.Extensions |
    Where-Object { $_.Oid.Value -eq "2.5.29.37" } |
    ForEach-Object {
        ([System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]::new($_, $false)).EnhancedKeyUsages
    } |
    ForEach-Object { $_.Value })
if ($enhancedKeyUsage -notcontains $codeSigningOid) {
    throw "Le certificat fourni n'est pas destiné à la signature de code."
}

$setupSignature = Get-AuthenticodeSignature -FilePath $setupPath
if ($null -eq $setupSignature.SignerCertificate) {
    throw "setup.exe ne contient aucune signature vérifiable."
}
if ($setupSignature.SignerCertificate.Thumbprint -ne $certificate.Thumbprint) {
    throw "Le certificat fourni ne correspond pas à celui qui a signé setup.exe."
}

Write-Host ""
Write-Host "Doctracker - installation de test" -ForegroundColor Cyan
Write-Host "Editeur      : $($certificate.Subject)"
Write-Host "Empreinte    : $($certificate.Thumbprint)"
Write-Host "Valable jusqu: $($certificate.NotAfter.ToString('dd/MM/yyyy HH:mm'))"
Write-Host ""
Write-Host "Ce certificat auto-signe sera approuve uniquement pour votre compte Windows." -ForegroundColor Yellow
$confirmation = Read-Host "Continuer ? [O/N]"
if ($confirmation -notmatch "^(o|oui|y|yes)$") {
    Write-Host "Installation annulee."
    exit 1
}

$rootStore = $null
$publisherStore = $null
try {
    $rootStore = [System.Security.Cryptography.X509Certificates.X509Store]::new(
        [System.Security.Cryptography.X509Certificates.StoreName]::Root,
        [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $rootStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $rootStore.Add($certificate)

    $publisherStore = [System.Security.Cryptography.X509Certificates.X509Store]::new(
        [System.Security.Cryptography.X509Certificates.StoreName]::TrustedPublisher,
        [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $publisherStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $publisherStore.Add($certificate)
}
finally {
    Close-CertificateStore -Store $publisherStore
    Close-CertificateStore -Store $rootStore
}

$trustedRoot = Get-ChildItem "Cert:\CurrentUser\Root\$($certificate.Thumbprint)" -ErrorAction SilentlyContinue
$trustedPublisher = Get-ChildItem "Cert:\CurrentUser\TrustedPublisher\$($certificate.Thumbprint)" -ErrorAction SilentlyContinue
if ($null -eq $trustedRoot -or $null -eq $trustedPublisher) {
    throw "Windows n'a pas enregistre le certificat dans les deux magasins de confiance attendus."
}

Write-Host ""
Write-Host "Certificat approuve. Lancement de l'installateur..." -ForegroundColor Green
$process = Start-Process -FilePath $setupPath -WorkingDirectory $packageDirectory -Wait -PassThru
if ($process.ExitCode -ne 0) {
    throw "setup.exe s'est termine avec le code $($process.ExitCode)."
}

Write-Host ""
Write-Host "Installation terminee. Fermez puis rouvrez Excel." -ForegroundColor Green
