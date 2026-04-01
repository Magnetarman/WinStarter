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

# Controllo privilegi di amministratore e auto-elevazione UAC
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "I permessi di Amministratore sono richiesti. Tentativo di riavvio con privilegi elevati..."
    $params = @()
    if ($PSBoundParameters.ContainsKey('CountdownSeconds')) { $params += "-CountdownSeconds $CountdownSeconds" }
    if ($PSBoundParameters.ContainsKey('SuppressIndividualReboot') -and $SuppressIndividualReboot) { $params += "-SuppressIndividualReboot" }
    
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $($params -join ' ')"
    Start-Process powershell.exe -ArgumentList $argList -Verb RunAs
    exit 0
}

$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "Set RustDesk By MagnetarMan"
$ToolkitVersion = "1.1.1"

$UserPath = "$env:APPDATA\RustDesk\config"
$SystemPath = "C:\Windows\ServiceProfiles\LocalService\AppData\Roaming\RustDesk\config"
$RustDeskInstaller = "$env:LOCALAPPDATA\WinToolkit\rustdesk\rustdesk-installer.msi"
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
    Write-StyledMessage Info "🛑 Arresto forzato componenti RustDesk..."
    try {
        # Arresto del servizio (gestisce errore se non esiste)
        Stop-Service -Name "RustDesk" -Force -ErrorAction SilentlyContinue
        
        # Kill dei processi attivi
        $null = taskkill /f /im rustdesk.exe 2>$null
        
        Write-StyledMessage Success "Servizi e processi RustDesk terminati con successo"
    }
    catch {
        Write-StyledMessage Warning "Nota: Alcuni componenti erano già inattivi"
    }
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
        $installArgs = "/i", "`"$InstallerPath`"", "REINSTALLMODE=amus", "/quiet", "/norestart"
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

function Inject-Configs {
    Write-StyledMessage Info "📥 Iniezione file di configurazione differenziata..."
    
    $ConfigMapping = @(
        # Mapper per UserPath
        @{ 
            Path = $UserPath 
            Files = @{
                "RustDesk2.toml"       = "https://raw.githubusercontent.com/Magnetarman/WinStarter/refs/heads/main/Asset/RustDesk/Roaming/RustDesk2.toml"
                "RustDesk_default.toml" = "https://raw.githubusercontent.com/Magnetarman/WinStarter/refs/heads/main/Asset/RustDesk/Roaming/RustDesk_default.toml"
                "RustDesk_local.toml"   = "https://raw.githubusercontent.com/Magnetarman/WinStarter/refs/heads/main/Asset/RustDesk/Roaming/RustDesk_local.toml"
            }
        },
        # Mapper per SystemPath
        @{ 
            Path = $SystemPath 
            Files = @{
                "RustDesk2.toml"       = "https://raw.githubusercontent.com/Magnetarman/WinStarter/refs/heads/main/Asset/RustDesk/GeneralConfig/RustDesk2.toml"
                "RustDesk_hwcodec.toml" = "https://raw.githubusercontent.com/Magnetarman/WinStarter/refs/heads/main/Asset/RustDesk/GeneralConfig/RustDesk_hwcodec.toml"
                "RustDesk_local.toml"   = "https://raw.githubusercontent.com/Magnetarman/WinStarter/refs/heads/main/Asset/RustDesk/Roaming/RustDesk_local.toml"
            }
        }
    )

    foreach ($map in $ConfigMapping) {
        $targetDir = $map.Path
        try {
            if (-not (Test-Path $targetDir)) { 
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null 
                Write-StyledMessage Info "Creazione directory: $targetDir"
            }

            foreach ($fileName in $map.Files.Keys) {
                $url = $map.Files[$fileName]
                $filePath = Join-Path $targetDir $fileName
                Write-StyledMessage Progress "Download $fileName -> $(Split-Path $targetDir -Leaf)"
                Invoke-WebRequest -Uri $url -OutFile $filePath -UseBasicParsing -ErrorAction Stop
            }
        }
        catch {
            Write-StyledMessage Error "Errore durante l'iniezione in $targetDir : $($_.Exception.Message)"
        }
    }
    Write-StyledMessage Success "Iniezione configurazioni completata"
}

function Initialize-RustDeskFirstRun {
    Write-StyledMessage Info "Inizializzazione database locale RustDesk (Step B)..."
    try {
        $rustDeskExe = "$env:ProgramFiles\RustDesk\rustdesk.exe"
        if (-not (Test-Path $rustDeskExe)) {
            Write-StyledMessage Warning "Eseguibile rustdesk.exe non trovato, il Double-Pass potrebbe fallire"
            return
        }
        # Avvia minimizzato/nascosto per generare config
        Start-Process -FilePath $rustDeskExe -WindowStyle Hidden -ErrorAction SilentlyContinue 
        Write-StyledMessage Progress "Attesa 10 secondi per l'inizializzazione..."
        Start-Sleep -Seconds 10
    }
    catch { Write-StyledMessage Warning "Errore durante l'inizializzazione: $($_.Exception.Message)" }
}

function Apply-ACL {
    Write-StyledMessage Info "🛡️ Applicazione ACL sui profili di sistema..."
    try {
        # Garantiamo che LOCAL SERVICE abbia FullControl sulla cartella di sistema
        if (Test-Path $SystemPath) {
            $acl = Get-Acl $SystemPath
            # Utilizziamo il SID S-1-5-19 (LOCAL SERVICE) per compatibilità universale (anche su OS Italiani)
            $sid = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-19")
            $permission = $sid, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"
            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($permission)
            $acl.SetAccessRule($accessRule)
            Set-Acl -Path $SystemPath -AclObject $acl -ErrorAction Stop
            Write-StyledMessage Success "Permessi LOCAL SERVICE applicati correttamente"
        }
    }
    catch {
        Write-StyledMessage Warning "Impossibile applicare ACL di sistema: $($_.Exception.Message)"
    }
}

function Start-RustDeskService {
    Write-StyledMessage Info "Avvio e impostazione servizio RustDesk su Automatico..."
    try {
        $service = Get-Service -Name "RustDesk" -ErrorAction SilentlyContinue
        if ($service) {
            Set-Service -Name "RustDesk" -StartupType Automatic
            Start-Service -Name "RustDesk" -ErrorAction SilentlyContinue
            Write-StyledMessage Success "Servizio RustDesk avviato su Automatico"
        }
    } catch { Write-StyledMessage Warning "Errore avvio servizio: $($_.Exception.Message)" }
}

Start-ToolkitLog -ToolName "SetRustDeskByMagnetarMan"
Show-Header -SubTitle "Set RustDesk By MagnetarMan"

Write-StyledMessage Info "🚀 AVVIO CONFIGURAZIONE RUSTDESK"

try {
    Write-StyledMessage Info "📋 [Step 0] Arresto preventivo servizi e processi RustDesk"
    Stop-RustDeskComponents

    if (-not (Download-RustDeskInstaller -DownloadPath $RustDeskInstaller)) {
        Write-StyledMessage Error "Impossibile procedere senza l'installer"
        exit 1
    }
    
    Write-StyledMessage Info "📋 [Step A] Gestione Installazione e Update"
    if (-not (Install-RustDesk -InstallerPath $RustDeskInstaller)) {
        Write-StyledMessage Error "Errore durante l'installazione"
        exit 1
    }

    Write-StyledMessage Info "📋 [Step B] Inizializzazione Primo Avvio (Generazione ID)"
    Initialize-RustDeskFirstRun

    Write-StyledMessage Info "📋 [Step C] Stop finale componenti e sblocco file"
    Stop-RustDeskComponents

    Write-StyledMessage Info "📋 [Step D] Iniezione Config Self-Hosted"
    Inject-Configs

    Write-StyledMessage Info "📋 [Step E] Applicazione Permessi e ACL"
    Apply-ACL

    Write-StyledMessage Info "📋 [Step F] Riavvio servizio RustDesk"
    Start-RustDeskService

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
