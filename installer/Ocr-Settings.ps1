# Shared by installation and Windows tests; loading this script has no side effect.
function Get-DoctrackerOcrAdvice {
    param([int]$LogicalProcessors, [double]$MemoryGb)
    $maximum=[Math]::Max(1,[Math]::Min(16,$LogicalProcessors))
    $cpu=[Math]::Max(1,[Math]::Floor($LogicalProcessors/2))
    $memory=if($MemoryGb -le 0){2}elseif($MemoryGb -lt 6){1}elseif($MemoryGb -lt 12){2}else{4}
    return @{Maximum=$maximum;Recommended=[int][Math]::Min($maximum,[Math]::Min($cpu,$memory))}
}
function Request-DoctrackerOcrWorkers {
    $logical=[Environment]::ProcessorCount
    $memory=0
    try {$memory=[double](Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory/1GB}catch{}
    $advice=Get-DoctrackerOcrAdvice $logical $memory
    $path=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Doctracker\ocr-workers.txt'
    $choice=$advice.Recommended
    if(Test-Path -LiteralPath $path){$existing=0;if([int]::TryParse((Get-Content -LiteralPath $path -Raw).Trim(),[ref]$existing) -and $existing -ge 1){$choice=[Math]::Min($advice.Maximum,$existing)}}
    Write-Host ""
    Write-Host "OCR local : $logical processeurs logiques ; $([Math]::Round($memory,1)) Go de RAM." -ForegroundColor Cyan
    Write-Host "Conseil : $($advice.Recommended) moteurs simultanes pour commencer. Maximum propose : $($advice.Maximum)."
    Write-Host "Davantage de moteurs peut ralentir Excel ou chauffer le PC. Ce maximum n'est pas une recommandation."
    Write-Host "Le reglage reste modifiable dans Doctracker > Reglages OCR."
    do {
        $answer=Read-Host "Nombre de moteurs [1-$($advice.Maximum), Entree = $choice]"
        if([string]::IsNullOrWhiteSpace($answer)){return [int]$choice}
        $parsed=0
        if([int]::TryParse($answer,[ref]$parsed) -and $parsed -ge 1 -and $parsed -le $advice.Maximum){return $parsed}
        Write-Host "Saisissez un entier entre 1 et $($advice.Maximum)." -ForegroundColor Yellow
    } while($true)
}
function Save-DoctrackerOcrWorkers {
    param([int]$Count)
    $directory=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Doctracker'
    [void][IO.Directory]::CreateDirectory($directory)
    $path=Join-Path $directory 'ocr-workers.txt'
    $temporary=$path+'.'+[Guid]::NewGuid().ToString('N')+'.tmp'
    try {
        [IO.File]::WriteAllText($temporary,$Count.ToString())
        # Windows PowerShell converts $null to an empty string for a .NET string
        # parameter. File.Replace requires a true null when no backup is requested.
        if([IO.File]::Exists($path)){[IO.File]::Replace($temporary,$path,[NullString]::Value)}else{[IO.File]::Move($temporary,$path)}
    } finally {if([IO.File]::Exists($temporary)){[IO.File]::Delete($temporary)}}
}
