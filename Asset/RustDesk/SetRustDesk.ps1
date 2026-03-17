<#
.SYNOPSIS
    Set RustDesk By MagnetarMan - Configura ed installa RustDesk con configurazioni personalizzate.

.DESCRIPTION
    Script autonomo per installare RustDesk e applicare configurazioni personalizzate.
    Scarica i file di configurazione da repository GitHub e riavvia il sistema.

.PARAMETER CountdownSeconds
    Numero di secondi per il countdown prima del riavvio.

.PARAMETER SuppressIndividualReboot
    Se specificato, sopprime il riavvio individuale.

#>

param(
    [int]$CountdownSeconds = 30,
    [switch]$SuppressIndividualReboot
)

$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "Set RustDesk By MagnetarMan"
$ToolkitVersion = "1.0.2"

$RustDeskConfig = "$env:APPDATA\RustDesk\config"
$RustDeskInstaller = "$env:LOCALAPPDATA\WinToolkit\rustdesk\rustdesk-installer.msi"
$RustDeskConfigPath = "$env:APPDATA\RustDesk\config"
$RustDeskReleaseAPI = "https://api.github.com/repos/rustdesk/rustdesk/releases/latest"

$Global:MsgStyles = @{
    Success  = @{ Icon = '✅'; Color = 'Green' }
    Warning  = @{ Icon = '⚠️'; Color = 'Yellow' }
    Error    = @{ Icon = '❌'; Color = 'Red' }
    Info     = @{ Icon = '💎'; Color = 'Cyan' }
    Progress = @{ Icon = '🔄'; Color = 'Magenta' }
}

$Global:Spinners = '⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏'.ToCharArray()
$Global:CurrentLogFile = $null
$Global:NeedsFinalReboot = $false

function Center-Text {
    param([string]$Text, [int]$Width = $Host.UI.RawUI.BufferSize.Width)
    $padding = [Math]::Max(0, [Math]::Floor(($Width - $Text.Length) / 2))
    return (' ' * $padding + $Text)
}

function Write-StyledMessage {
    param([ValidateSet('Success', 'Warning', 'Error', 'Info', 'Progress')]$Type, [string]$Text)
    $style = $Global:MsgStyles[$Type]
    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host "[$timestamp] $($style.Icon) $Text" -ForegroundColor $style.Color
    
    if ($Global:CurrentLogFile) {
        $logLevel = switch ($Type) { 'Success' { 'SUCCESS' } 'Warning' { 'WARNING' } 'Error' { 'ERROR' } default { 'INFO' } }
        $line = "[$timestamp] [$logLevel] $Text"
        try { Add-Content -Path $Global:CurrentLogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
    }
}

function Show-Header {
    param([string]$SubTitle = "Menu")
    Clear-Host
    $width = $Host.UI.RawUI.BufferSize.Width
    $asciiArt = @(
        '      __        __  _   _   _ ',
        '      \ \      / / | | | \ | |',
        '       \ \ /\ / /  | | |  \| |',
        '        \ V  V /   | | | |\  |',
        '         \_/\_/    |_| |_| \_|',
        '',
        "       Set RustDesk By MagnetarMan",
        "       Versione $ToolkitVersion"
    )
    Write-Host ('═' * ($width - 1)) -ForegroundColor Green
    foreach ($line in $asciiArt) { Write-Host (Center-Text $line $width) -ForegroundColor White }
    Write-Host ('═' * ($width - 1)) -ForegroundColor Green
    Write-Host ''
}

function Start-ToolkitLog {
    param([string]$ToolName)
    try { Stop-Transcript -ErrorAction SilentlyContinue } catch {}
    $dateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $logdir = "$env:LOCALAPPDATA\WinToolkit\logs"
    if (-not (Test-Path $logdir)) { New-Item -Path $logdir -ItemType Directory -Force | Out-Null }
    $Global:CurrentLogFile = "$logdir\${ToolName}_$dateTime.log"
    
    $header = @"
[START LOG HEADER]
ToolName    : $ToolName
Start time  : $dateTime
Username    : $([Environment]::UserDomainName)\$([Environment]::UserName)
[END LOG HEADER]
"@
    try { Add-Content -Path $Global:CurrentLogFile -Value $header -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

function Start-InterruptibleCountdown {
    param([int]$Seconds = 30, [string]$Message = "Riavvio automatico", [switch]$Suppress)
    if ($Suppress) { return $true }
    Write-StyledMessage -Type 'Info' -Text '💡 Premi un tasto qualsiasi per annullare...'
    Write-Host ''
    for ($i = $Seconds; $i -gt 0; $i--) {
        if ([Console]::KeyAvailable) {
            $null = [Console]::ReadKey($true)
            Write-Host "`n"
            Write-StyledMessage -Type 'Warning' -Text '⏸️ Riavvio del sistema annullato.'
            return $false
        }
        $percent = [Math]::Round((($Seconds - $i) / $Seconds) * 100)
        $filled = [Math]::Floor($percent * 20 / 100)
        $remaining = 20 - $filled
        $bar = "[$('█' * $filled)$('▒' * $remaining)]"
        Write-Host "`r⏰ $Message tra $i secondi $bar" -NoNewline -ForegroundColor Red
        Start-Sleep 1
    }
    Write-Host "`n"
    return $true
}

function Stop-RustDeskComponents {
    $servicesFound = $false
    foreach ($service in @("RustDesk", "rustdesk")) {
        $serviceObj = Get-Service -Name $service -ErrorAction SilentlyContinue
        if ($serviceObj) {
            Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
            $servicesFound = $true
        }
    }
    if ($servicesFound) { Write-StyledMessage Success "Servizi RustDesk arrestati" }

    $processesFound = $false
    foreach ($process in @("rustdesk", "RustDesk")) {
        $runningProcesses = Get-Process -Name $process -ErrorAction SilentlyContinue
        if ($runningProcesses) {
            $runningProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
            $processesFound = $true
        }
    }
    if ($processesFound) { Write-StyledMessage Success "Processi RustDesk terminati" }
    if (-not $servicesFound -and -not $processesFound) { Write-StyledMessage Warning "Nessun componente RustDesk attivo trovato" }
    Start-Sleep 2
}

function Get-LatestRustDeskRelease {
    try {
        $response = Invoke-RestMethod -Uri $RustDeskReleaseAPI -Method Get -ErrorAction Stop
        $msiAsset = $response.assets | Where-Object { $_.name -like "rustdesk-*-x86_64.msi" } | Select-Object -First 1
        if ($msiAsset) {
            return @{ Version = $response.tag_name; DownloadUrl = $msiAsset.browser_download_url; FileName = $msiAsset.name }
        }
        Write-StyledMessage -Type 'Error' -Text "Nessun installer .msi trovato nella release"
        return $null
    }
    catch {
        Write-StyledMessage -Type 'Error' -Text "Errore connessione GitHub API: $($_.Exception.Message)"
        return $null
    }
}

function Download-RustDeskInstaller {
    param([string]$DownloadPath)
    Write-StyledMessage -Type 'Info' -Text "Download installer RustDesk in corso..."
    $releaseInfo = Get-LatestRustDeskRelease
    if (-not $releaseInfo) { return $false }
    Write-StyledMessage -Type 'Info' -Text "📥 Versione rilevata: $($releaseInfo.Version)"
    $parentDir = Split-Path $DownloadPath -Parent
    try {
        if (-not (Test-Path $parentDir)) { New-Item -ItemType Directory -Path $parentDir -Force | Out-Null }
        if (Test-Path $DownloadPath) { Remove-Item $DownloadPath -Force -ErrorAction Stop }
        Invoke-WebRequest -Uri $releaseInfo.DownloadUrl -OutFile $DownloadPath -UseBasicParsing -ErrorAction Stop
        if (Test-Path $DownloadPath) {
            Write-StyledMessage -Type 'Success' -Text "Installer $($releaseInfo.FileName) scaricato con successo"
            return $true
        }
    }
    catch { Write-StyledMessage -Type 'Error' -Text "Errore download: $($_.Exception.Message)" }
    return $false
}

function Install-RustDesk {
    param([string]$InstallerPath)
    Write-StyledMessage -Type 'Info' -Text "Installazione RustDesk"
    try {
        $installArgs = "/i", "`"$InstallerPath`"", "/quiet", "/norestart"
        $process = Start-Process "msiexec.exe" -ArgumentList $installArgs -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
        Start-Sleep 10
        if ($process.ExitCode -eq 0) {
            Write-StyledMessage -Type 'Success' -Text "RustDesk installato"
            return $true
        }
        else { Write-StyledMessage -Type 'Error' -Text "Errore installazione (Exit Code: $($process.ExitCode))" }
    }
    catch { Write-StyledMessage -Type 'Error' -Text "Errore durante installazione: $($_.Exception.Message)" }
    return $false
}

function Clear-RustDeskConfig {
    Write-StyledMessage Info "Pulizia configurazioni esistenti..."
    $configDir = "$RustDeskConfigPath\config"
    try {
        if (-not (Test-Path $RustDeskConfigPath)) {
            New-Item -ItemType Directory -Path $RustDeskConfigPath -Force | Out-Null
            Write-StyledMessage Info "Cartella RustDesk creata"
        }
        if (Test-Path $configDir) {
            Remove-Item $configDir -Recurse -Force -ErrorAction Stop
            Write-StyledMessage Success "Cartella config eliminata"
            Start-Sleep 1
        }
        else { Write-StyledMessage Warning "Cartella config non trovata" }
    }
    catch { Write-StyledMessage Error "Errore pulizia config: $($_.Exception.Message)" }
}

function Download-RustDeskConfigFiles {
    Write-StyledMessage Info "Download file di configurazione..."
    $configDir = "$env:APPDATA\RustDesk\config"
    try {
        if (-not (Test-Path $configDir)) { New-Item -ItemType Directory -Path $configDir -Force | Out-Null }
        $configUrls = @{
            "RustDesk.toml"       = "https://raw.githubusercontent.com/Magnetarman/WinStarter/refs/heads/main/Asset/RustDesk/RustDesk.toml"
            "RustDesk_local.toml" = "https://raw.githubusercontent.com/Magnetarman/WinStarter/refs/heads/main/Asset/RustDesk/RustDesk_local.toml"
            "RustDesk2.toml"      = "https://raw.githubusercontent.com/Magnetarman/WinStarter/refs/heads/main/Asset/RustDesk/RustDesk2.toml"
        }
        $downloaded = 0
        foreach ($fileName in $configUrls.Keys) {
            $url = $configUrls[$fileName]
            $filePath = Join-Path $configDir $fileName
            try {
                Invoke-WebRequest -Uri $url -OutFile $filePath -UseBasicParsing -ErrorAction Stop
                $downloaded++
            }
            catch { Write-StyledMessage Error "Errore download $fileName`: $($_.Exception.Message)" }
        }
        if ($downloaded -eq $configUrls.Count) { Write-StyledMessage Success "Tutti i file di configurazione scaricati ($downloaded/$($configUrls.Count))" }
        else { Write-StyledMessage Warning "Scaricati $downloaded/$($configUrls.Count) file di configurazione" }
    }
    catch { Write-StyledMessage Error "Errore durante download configurazioni: $($_.Exception.Message)" }
}

Start-ToolkitLog -ToolName "SetRustDeskByMagnetarMan"
Show-Header -SubTitle "Set RustDesk By MagnetarMan"

Write-StyledMessage Info "🚀 AVVIO CONFIGURAZIONE RUSTDESK"

try {
    if (-not (Download-RustDeskInstaller -DownloadPath $RustDeskInstaller)) {
        Write-StyledMessage Error "Impossibile procedere senza l'installer"
        exit 1
    }
    if (-not (Install-RustDesk -InstallerPath $RustDeskInstaller)) {
        Write-StyledMessage Error "Errore durante l'installazione"
        exit 1
    }

    Write-StyledMessage Info "📋 Arresto servizi e processi RustDesk"
    Stop-RustDeskComponents

    Write-StyledMessage Info "📋 Pulizia configurazioni"
    Clear-RustDeskConfig

    Write-StyledMessage Info "📋 Download configurazioni"
    Download-RustDeskConfigFiles

    Write-Host ""
    Write-StyledMessage Success "🎉 CONFIGURAZIONE RUSTDESK COMPLETATA"
    Write-StyledMessage Info "🔄 Per applicare le modifiche il PC verrà riavviato"

    if ($SuppressIndividualReboot) {
        $Global:NeedsFinalReboot = $true
        Write-StyledMessage -Type 'Info' -Text "❌ Riavvio individuale soppresso. Verrà gestito un riavvio finale."
    }
    else {
        $shouldReboot = Start-InterruptibleCountdown -Seconds $CountdownSeconds -Message "Per applicare le modifiche è necessario riavviare il sistema"
        if ($shouldReboot) { Restart-Computer -Force }
        else { Write-StyledMessage Success "🎉 Configurazione RustDesk completata (riavvio annullato)" }
    }
}
catch {
    Write-StyledMessage Error "ERRORE CRITICO: $($_.Exception.Message)"
    Write-StyledMessage Info "💡 Verifica connessione Internet e riprova"
    exit 1
}
finally {
    Write-StyledMessage Success "🎯 Setup RustDesk terminato"
    try { Stop-Transcript | Out-Null } catch {}
    exit 0
}
