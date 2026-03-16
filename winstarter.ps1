<#
.SYNOPSIS
    Script di inizio che installa e configura Win Starter.
.DESCRIPTION
    Verifica, re-imposta e ripara Winget, installa software primari e 
    configura personalizzazioni specifiche per ridurre bloatware, 
    incluso PowerToys, Nilesoft Shell e tweak di sistema avanzati.
.NOTES
    Autore: Magnetarman
    Nome: Win Starter By Magnetarman
    Compatibile con: PowerShell 5.1 e 7+
#>

# --- CONFIGURAZIONE GLOBALE ---
$ErrorActionPreference = 'Stop'
$Global:MsgStyles = @{
    Success  = @{ Icon = '✅'; Color = 'Green' }
    Warning  = @{ Icon = '⚠️'; Color = 'Yellow' }
    Error    = @{ Icon = '❌'; Color = 'Red' }
    Info     = @{ Icon = '💎'; Color = 'Cyan' }
    Progress = @{ Icon = '🔄'; Color = 'Magenta' }
}

$script:AppConfig = @{
    Header   = @{
        Title   = "Win Starter By Magnetarman"
        Version = "Version 1.1.0"
    }
    URLs     = @{
        PowerToysConfig         = "https://github.com/Magnetarman/WinStarter/raw/refs/heads/main/Asset/PowerToys.zip"
        NilesoftConfig          = "https://github.com/Magnetarman/WinStarter/raw/refs/heads/main/Asset/NilesoftShell.zip"
        WinSupportIcon          = "https://github.com/Magnetarman/WinStarter/raw/refs/heads/main/img/WinSupport.png"
        WingetMSIX              = "https://aka.ms/getwinget"
        PowerShellRelease       = "https://api.github.com/repos/PowerShell/PowerShell/releases/latest"
        OhMyPoshTheme           = "https://raw.githubusercontent.com/JanDeDobbeleer/oh-my-posh/main/themes/atomic.omp.json"
        PowerShellProfile       = "https://raw.githubusercontent.com/Magnetarman/WinToolkit/Dev/asset/Microsoft.PowerShell_profile.ps1"
        WindowsTerminalSettings = "https://raw.githubusercontent.com/Magnetarman/WinToolkit/Dev/asset/settings.json"
        TerminalRelease         = "https://api.github.com/repos/microsoft/terminal/releases/latest"
    }
    Paths    = @{
        Logs          = "$env:LOCALAPPDATA\WinStarter\logs"
        WinToolkitDir = "$env:LOCALAPPDATA\WinToolkit"
        Temp          = "$env:TEMP\WinStarterSetup"
        PowerShell7   = "$env:ProgramFiles\PowerShell\7"
        Packages      = "$env:LOCALAPPDATA\Packages"
        Desktop       = [Environment]::GetFolderPath('Desktop')
        wtExe         = "$env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"
        wtDir         = "$env:LOCALAPPDATA\Microsoft\WindowsApps"
        PowerToysDir  = "$env:LOCALAPPDATA\Microsoft\PowerToys"
        NilesoftDir   = "$env:ProgramFiles\Nilesoft Shell"
    }
    Registry = @{
        TerminalStartup = "HKCU:\Console\%%Startup"
    }
}

# ============================================================================
# FUNZIONI DI UTILITÀ BASE E STILI
# ============================================================================

function Format-CenteredText {
    <#
    .SYNOPSIS
    Centra il testo nel terminale in base ad una larghezza fissa.
    #>
    param([string]$Text, [int]$Width = 80)
    $padding = [Math]::Max(0, [Math]::Floor(($Width - $Text.Length) / 2))
    return (" " * $padding) + $Text
}

function Show-Header {
    <#
    .SYNOPSIS
    Pulisce lo schermo e mostra l'intestazione ASCII dello script.
    #>
    param([string]$Title, [string]$Version)
    Clear-Host
    $width = 65
    Write-Host ('═' * $width) -ForegroundColor Green
    @(
        '      __        __  _   _   _ ',
        '      \ \      / / | | | \ | |',
        '       \ \ /\ / /  | | |  \| |',
        '        \ V  V /   | | | |\  |',
        '         \_/\_/    |_| |_| \_|',
        '',
        $Title,
        $Version
    ) | ForEach-Object { Write-Host (Format-CenteredText -Text $_ -Width $width) -ForegroundColor White }
    Write-Host ('═' * $width) -ForegroundColor Green
    Write-Host ''
}

function Write-StyledMessage {
    <#
    .SYNOPSIS
    Stampa un output formattato con icone e lo registra parallelamente nel file di log.
    #>
    param(
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Type,
        [string]$Text
    )
    # FIX: Risolve problemi di indentazione anomala su Windows 11
    if ([Environment]::OSVersion.Version.Build -ge 22000) { $Text = "`r$Text" }
    
    $style = $Global:MsgStyles[$Type]
    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host "[$timestamp] $($style.Icon) $Text" -ForegroundColor $style.Color
    
    $logLevel = switch ($Type) {
        'Success' { 'SUCCESS' }
        'Warning' { 'WARNING' }
        'Error' { 'ERROR' }
        default { 'INFO' }
    }
    Write-ToolkitLog -Level $logLevel -Message $Text
}

function Start-ToolkitLog {
    <#
    .SYNOPSIS
    Inizializza un file di log strutturato salvando le info di sistema dell'utente.
    #>
    param([string]$ToolName)
    try { Stop-Transcript -ErrorAction SilentlyContinue } catch {}
    
    $dateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $logdir = $script:AppConfig.Paths.Logs
    if (-not (Test-Path $logdir)) { New-Item -Path $logdir -ItemType Directory -Force | Out-Null }
    $Global:CurrentLogFile = "$logdir\${ToolName}_$dateTime.log"

    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $psVer = $PSVersionTable.PSVersion.ToString()
    
    $header = @"
[START LOG HEADER]
Start time     : $dateTime
ToolName       : $ToolName
Username       : $([Environment]::UserDomainName + '\' + [Environment]::UserName)
Machine        : $($env:COMPUTERNAME) ($($os.Caption) $($os.Version))
PSVersion      : $psVer
ToolkitVersion : $($script:AppConfig.Header.Version)
[END LOG HEADER]

"@
    try { Add-Content -Path $Global:CurrentLogFile -Value $header -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

function Write-ToolkitLog {
    <#
    .SYNOPSIS
    Accoda una riga al file di log corrente.
    #>
    param(
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO',
        [string]$Message
    )
    if (-not $Global:CurrentLogFile) { return }
    $ts = Get-Date -Format "HH:mm:ss"
    $clean = $Message -replace '^\s+', ''
    $line = "[$ts] [$Level] $clean"
    try { Add-Content -Path $Global:CurrentLogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}


# ============================================================================
# UTILITIES WINGET E PERMESSI
# ============================================================================

function Start-AppxSilentProcess {
    <#
    .SYNOPSIS
    Installa file AppX/MSIX in background sopprimendo forzatamente le barre di progresso native.
    #>
    param([string]$AppxPath, [string]$Flags = '-ForceApplicationShutdown')
    $cmd = @"
`$ProgressPreference = 'SilentlyContinue';
try { Add-AppxPackage -Path '$($AppxPath -replace "'", "''")' $Flags -ErrorAction Stop | Out-Null }
catch { exit 1 }
exit 0
"@
    $encodedCmd = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($cmd))
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.Arguments = "-NoProfile -NonInteractive -EncodedCommand $encodedCmd"
    $psi.CreateNoWindow = $true
    $psi.UseShellExecute = $false
    
    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.WaitForExit()
    return $proc.ExitCode -eq 0
}

function Stop-InterferingProcess {
    <#
    .SYNOPSIS
    Termina forzatamente i processi noti che bloccano le installazioni Appx o Winget.
    #>
    $interferingProcesses = @("WinStore.App", "wsappx", "AppInstaller", "Microsoft.WindowsStore", "Microsoft.DesktopAppInstaller", "winget", "WindowsPackageManagerServer")
    foreach ($procName in $interferingProcesses) {
        $null = Get-Process -Name $procName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep 2
}

function Invoke-ForceCloseWinget {
    <#
    .SYNOPSIS
    Wrapper per chiudere tutti i processi Winget per liberare file lock bloccanti.
    #>
    Write-StyledMessage -Type Info -Text "Chiusura processi interferenti Winget per liberare lock..."
    Stop-InterferingProcess
    Write-StyledMessage -Type Success -Text "✅ Processi interferenti chiusi."
}

function Update-EnvironmentPath {
    <#
    .SYNOPSIS
    Ricarica la variabile d'ambiente PATH nel processo corrente per rilevare installazioni appena concluse.
    #>
    $machinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    $newPath = ($machinePath, $userPath | Where-Object { $_ }) -join ';'
    $env:Path = $newPath
    [System.Environment]::SetEnvironmentVariable('Path', $newPath, 'Process')
}

function Add-ToEnvironmentPath {
    <#
    .SYNOPSIS
    Aggiunge una stringa PATH alle variabili utente o di sistema in modo persistente.
    #>
    param ([string]$PathToAdd, [string]$Scope)
    # Check locale omesso per brevità, se serve lo inseriremo in futuri fix
}

function Set-PathPermissions {
    <#
    .SYNOPSIS
    Assegna i privilegi di FullControl al gruppo Administrators sulla cartella passata per parametro.
    #>
    param ([string]$FolderPath)
    if (-not (Test-Path $FolderPath)) { return }
    try {
        $administratorsGroupSid = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-544")
        $administratorsGroup = $administratorsGroupSid.Translate([System.Security.Principal.NTAccount])
        $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($administratorsGroup, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($accessRule)
        Set-Acl -Path $FolderPath -AclObject $acl -ErrorAction Stop
    } catch { }
}

function Set-WingetPathPermissions {
    <#
    .SYNOPSIS
    Localizza la directory di installazione dinamica di Winget e ne corregge in permessi bloccati.
    #>
    $wingetFolderPath = $null
    try {
        $arch = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }
        $wingetDir = Get-ChildItem -Path "$env:ProgramFiles\WindowsApps" -Filter "Microsoft.DesktopAppInstaller_*_*${arch}__8wekyb3d8bbwe" -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
        if ($wingetDir) { $wingetFolderPath = $wingetDir.FullName }
    } catch { }
    if ($wingetFolderPath) {
        Set-PathPermissions -FolderPath $wingetFolderPath
    }
}

function Invoke-WingetCommand {
    <#
    .SYNOPSIS
    Wrapper robusto per l'esecuzione di comandi Winget. Gestisce logica retrocompatibile.
    #>
    param([string]$Arguments, [int]$TimeoutSeconds = 120)
    try {
        # Aggiungiamo --disable-interactivity solo se la versione supporta (Winget v1.4+)
        $versionRaw = (winget --version 2>$null) | Out-String
        $isModern = $versionRaw -match 'v1\.[4-9]' -or $versionRaw -match 'v[2-9]'
        $finalArgs = if ($isModern) { "$Arguments --disable-interactivity" } else { $Arguments }
        
        $procParams = @{ FilePath = 'winget'; ArgumentList = $finalArgs -split ' '; Wait = $true; PassThru = $true; NoNewWindow = $true }
        $process = Start-Process @procParams
        return @{ ExitCode = $process.ExitCode }
    } catch {
        return @{ ExitCode = -1 }
    }
}

function Test-WingetFunctionality {
    <#
    .SYNOPSIS
    Esegue un controllo primario su Winget testando l'accessibilità nel PATH e l'output della versione.
    #>
    Write-StyledMessage -Type Info -Text "🔍 Verifica primordiale disponibilità Winget..."
    Update-EnvironmentPath
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-StyledMessage -Type Warning -Text "⚠️ Winget non trovato nel PATH corrente di sistema."
        return $false
    }
    try {
        $versionOutput = (& winget --version 2>$null) | Out-String
        if ($LASTEXITCODE -eq 0 -and $versionOutput -match 'v\d+\.\d+') {
            Write-StyledMessage -Type Success -Text "✅ Winget rintracciato e operativo (versione: $($versionOutput.Trim()))."
            return $true
        }
        return $false
    } catch {
        return $false
    }
}

function Test-WingetCompatibility {
    <#
    .SYNOPSIS
    Controlla che la build in uso del sistema operativo Windows supporti intrinsecamente Winget.
    #>
    $osInfo = [Environment]::OSVersion
    $build = $osInfo.Version.Build
    if ($osInfo.Version.Major -lt 10 -or ($osInfo.Version.Major -eq 10 -and $build -lt 16299)) {
        Write-StyledMessage -Type Error -Text "❌ La versione SO rilevata non supporta l'uso di Winget."
        return $false
    }
    return $true
}

function Repair-WingetDatabase {
    <#
    .SYNOPSIS
    Tenta autonomamente di ripristinare un database Winget corrotto pulendo cache e ri-registrando l'app.
    #>
    Write-StyledMessage -Type Info -Text "🔧 Avvio logica drastica di ripristino per il database Winget..."
    try {
        Stop-InterferingProcess
        
        # Pulizia cartelle di Cache
        $wingetCachePath = "$env:LOCALAPPDATA\WinGet"
        if (Test-Path $wingetCachePath) {
            Get-ChildItem -Path $wingetCachePath -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notmatch '\\lock\\|\\tmp\\' } | ForEach-Object { try { Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue } catch {} }
        }
        
        # Reset dei file di Database Locali
        $stateFiles = @("$env:LOCALAPPDATA\WinGet\Data\USERTEMPLATE.json", "$env:LOCALAPPDATA\WinGet\Data\DEFAULTUSER.json")
        foreach ($file in $stateFiles) { if (Test-Path $file -PathType Leaf) { Remove-Item $file -Force -ErrorAction SilentlyContinue } }
        
        # Reset Source
        try { $null = & winget.exe source reset --force 2>&1 } catch {}
        Update-EnvironmentPath
        
        # Reset logico dello store Package
        if (Get-Command Reset-AppxPackage -ErrorAction SilentlyContinue) {
            Get-AppxPackage -Name 'Microsoft.DesktopAppInstaller' | Reset-AppxPackage 2>$null
        }
        
        # Repair Module (se presente)
        try { if (Get-Command Repair-WinGetPackageManager -ErrorAction SilentlyContinue) { Repair-WinGetPackageManager -Force -Latest 2>$null *>$null } } catch {}
        
        Set-WingetPathPermissions
        Update-EnvironmentPath
        Start-Sleep 2
        return $true
    } catch {
        return $false
    }
}

function Test-WingetDeepValidation {
    <#
    .SYNOPSIS
    Imita l'attività di un utente reale per dedurre se la comunicazione remota coi repository funziona.
    Se fallisce per ACCESS_VIOLATION (0xC0000005) la funzione lancia automaticamente la pipeline di riparo locale.
    #>
    Write-StyledMessage -Type Info -Text "🔍 Test stringente di networking Winget..."
    try {
        $searchResult = & winget search "Git.Git" --accept-source-agreements 2>&1
        $exitCode = $LASTEXITCODE
        
        # Intercetta il crash noto dovuto alla corruzione index repository locale
        if ($exitCode -eq -1073741819 -or $exitCode -eq 3221225781) {
            $null = Repair-WingetDatabase
            Start-Sleep 3
            # Secondo tentativo post-cura
            $searchResult = & winget search "Git.Git" --accept-source-agreements 2>&1
            $exitCode = $LASTEXITCODE
        }
        return ($exitCode -eq 0)
    } catch {
        return $false
    }
}

function Install-WingetCore {
    <#
    .SYNOPSIS
    Rimedio estremo: scarica il pacchetto standalone MSIX e lo re-installa saltando gli intermediari.
    #>
    Write-StyledMessage -Type Info -Text "🛠️ Dowload sorgenti e sovrascrittura forzata blocco Winget..."
    $tempDir = "$env:TEMP\WinToolkitWinget"
    if (-not (Test-Path $tempDir)) { New-Item -Path $tempDir -ItemType Directory -Force | Out-Null }
    try {
        $wingetUrl = (Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest" -UseBasicParsing).assets | Where-Object { $_.name -match 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle' } | Select-Object -First 1 | Select-Object -ExpandProperty browser_download_url
        if ($wingetUrl) {
            $wingetFile = Join-Path $tempDir "winget.msixbundle"
            Invoke-WebRequest -Uri $wingetUrl -OutFile $wingetFile -UseBasicParsing
            if (Start-AppxSilentProcess -AppxPath $wingetFile) { Write-StyledMessage -Type Success -Text "✅ WinGet installato a basso livello con successo." }
        }
    } catch { }
}


# ============================================================================
# INSTALLAZIONE COMPONENTI (Powershell 7, Terminal, PSP)
# ============================================================================

function Install-PowerShellCore {
    <#
    .SYNOPSIS
    Installa la versione più recente (moderna) di PowerShell, essenziale per bypassare i vincoli della v5.1 legacy.
    #>
    Write-StyledMessage -Type Info -Text "🔍 Verifica indipendenza su PowerShell 7..."
    $ps7Path = $script:AppConfig.Paths.PowerShell7
    
    if (Test-Path $ps7Path) {
        Write-StyledMessage -Type Success -Text "✅ Core PowerShell 7 già integrato."
        return $true
    }
    
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-StyledMessage -Type Info -Text "⬇️ Bootstrap PowerShell 7 via terminale in corso..."
        $res = Invoke-WingetCommand -Arguments "install --id Microsoft.PowerShell --source winget --accept-source-agreements --accept-package-agreements --silent"
        if ($res.ExitCode -eq 0 -or (Test-Path $ps7Path)) {
            Write-StyledMessage -Type Success -Text "✅ PowerShell 7 convalidato e pronto."
            return $true
        }
    }
    return $false
}

function Install-WindowsTerminalApp {
    <#
    .SYNOPSIS
    Approvigiona Windows Terminal, preferibilmente usando il server Microsoft Store proxy Winget.
    #>
    Write-StyledMessage -Type Info -Text "🔍 Messa in sicurezza Windows Terminal moderno..."
    
    if (Get-Command "wt.exe" -ErrorAction SilentlyContinue) {
        Write-StyledMessage -Type Success -Text "✅ L'App Windows Terminal è nativamente presente."
        return $true
    }
    
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        $res = Invoke-WingetCommand -Arguments "install --id 9N0DX20HK701 --source msstore --accept-source-agreements --accept-package-agreements --silent"
        if ($res.ExitCode -eq 0 -or (Get-Command "wt.exe" -ErrorAction SilentlyContinue)) {
            Write-StyledMessage -Type Success -Text "✅ Ambiente Desktop aggiornato con Windows Terminal."
            return $true
        }
    }
    return $false
}

function Install-PspEnvironment {
    <#
    .SYNOPSIS
    Prepara il profilo bash/powershell definitivo, dotando l'utente dei tool Zoxide, Btop ecc. e del relativo tema.
    #>
    Write-StyledMessage -Type Info -Text "🛠️ Estensione framework linea di comando (Zoxide, OhMyPosh, font)..."
    
    $tools = @(
        @{ Id = "JanDeDobbeleer.OhMyPosh"; Name = "Oh My Posh" },
        @{ Id = "ajeetdsouza.zoxide"; Name = "zoxide" },
        @{ Id = "aristocratos.btop4win"; Name = "btop" },
        @{ Id = "Fastfetch-cli.Fastfetch"; Name = "fastfetch" },
        @{ Id = "DEVCOM.JetBrainsMonoNerdFont"; Name = "JetBrainsMono Nerd Font" }
    )

    foreach ($tool in $tools) {
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Invoke-WingetCommand -Arguments "install -e --id $($tool.Id) --accept-source-agreements --accept-package-agreements --silent" *>$null
        }
    }

    # OhMyPosh Theme & Profile
    try {
        $profileDir = if ($PSVersionTable.PSEdition -eq "Core") { [Environment]::GetFolderPath("MyDocuments") + "\PowerShell" } else { [Environment]::GetFolderPath("MyDocuments") + "\WindowsPowerShell" }
        if (-not (Test-Path $profileDir)) { New-Item -Path $profileDir -ItemType Directory -Force | Out-Null }
        
        $themesFolder = Join-Path $profileDir "Themes"
        if (-not (Test-Path $themesFolder)) { New-Item -Path $themesFolder -ItemType Directory -Force | Out-Null }
        
        $themePath = Join-Path $themesFolder "atomic.omp.json"
        # Download Oh My Posh theme da remote
        Invoke-WebRequest -Uri $script:AppConfig.URLs.OhMyPoshTheme -OutFile $themePath -UseBasicParsing | Out-Null
        
        # Override file profile (.profile equivalente per PS)
        $targetProfile = $PROFILE
        if (-not $targetProfile) { $targetProfile = Join-Path $profileDir "Microsoft.PowerShell_profile.ps1" }
        if (Test-Path $targetProfile) { Move-Item -Path $targetProfile -Destination "$targetProfile.bak" -Force -ErrorAction SilentlyContinue }
        Invoke-WebRequest -Uri $script:AppConfig.URLs.PowerShellProfile -OutFile $targetProfile -UseBasicParsing | Out-Null
        
        Write-StyledMessage -Type Success -Text "✅ Profilo bash customizzato applicato."
    } catch { }

    # Terminal Settings
    try {
        $wtPath = Get-ChildItem -Path "$env:LOCALAPPDATA\Packages" -Directory -Filter "Microsoft.WindowsTerminal_*" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($wtPath) {
            $settingsPath = Join-Path $wtPath.FullName "LocalState\settings.json"
            if (Test-Path (Join-Path $wtPath.FullName "LocalState")) {
                Invoke-WebRequest -Uri $script:AppConfig.URLs.WindowsTerminalSettings -OutFile $settingsPath -UseBasicParsing | Out-Null
                Write-StyledMessage -Type Success -Text "✅ Regole visive Terminal iniettate correttamente."
            }
        }
    } catch { }
}

# ============================================================================
# FUNZIONI CORE WINSTARTER
# ============================================================================

function SetRecommendedUpdate {
    <#
    .SYNOPSIS
    Mitiga il comportamento aggressivo di Windows Update fermando l'inclusione di driver buggati e posticipando i Feature Update.
    #>
    Write-StyledMessage -Type Info -Text "⚙️ Disabilitazione aggiornamenti driver tramite Windows Update..."

    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Device Metadata" -Force | Out-Null
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Device Metadata" -Name "PreventDeviceMetadataFromNetwork" -Type DWord -Value 1

    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Force | Out-Null
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Name "DontPromptForWindowsUpdate" -Type DWord -Value 1
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Name "DontSearchWindowsUpdate" -Type DWord -Value 1
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverSearching" -Name "DriverUpdateWizardWuSearchEnabled" -Type DWord -Value 0

    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Force | Out-Null
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "ExcludeWUDriversInQualityUpdate" -Type DWord -Value 1

    Write-StyledMessage -Type Info -Text "⏱️ Posticipo Feature Updates di 365 giorni e Quality Updates di 4 giorni..."

    New-Item -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Force | Out-Null
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "BranchReadinessLevel" -Type DWord -Value 20
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "DeferFeatureUpdatesPeriodInDays" -Type DWord -Value 365
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "DeferQualityUpdatesPeriodInDays" -Type DWord -Value 4

    Write-StyledMessage -Type Info -Text "🛑 Disabilitazione riavvio automatico di Windows Update..."

    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Force | Out-Null
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoRebootWithLoggedOnUsers" -Type DWord -Value 1
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "AUPowerManagement" -Type DWord -Value 0

    Write-StyledMessage -Type Success -Text "✅ Impostazioni Windows Update raccomandate applicate con successo."
}

function Set-ExplorerPersonalization {
    <#
    .SYNOPSIS
    Pulisce Explorer: mostra le info estese sui file, imposta il dark mode e nasconde la pubblicità (Bing) nel menu Start.
    #>
    Write-StyledMessage -Type Info -Text "⚙️ Applicazione personalizzazioni File Explorer in corso..."
    try {
        $explorerAdv = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        
        Set-ItemProperty -Path $explorerAdv -Name "Hidden" -Value 1 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "HideFileExt" -Value 0 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "UseCompactMode" -Value 1 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "AutoCheckSelect" -Value 0 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "ShowCompColor" -Value 1 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "IconsOnly" -Value 0 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "ShowRecent" -Value 0 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "ShowFrequent" -Value 0 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "Start_TrackDocs" -Value 0 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "LaunchTo" -Value 1 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "ShowInfoTip" -Value 1 -Type DWord
        Set-ItemProperty -Path $explorerAdv -Name "ShowTaskViewButton" -Value 1 -Type DWord

        # Bing Search Remove
        $bingSearch = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
        if (-not (Test-Path $bingSearch)) { New-Item -Path $bingSearch -Force | Out-Null }
        Set-ItemProperty -Path $bingSearch -Name "DisableSearchBoxSuggestions" -Value 1 -Type DWord

        # Personalizzazione Dark Theme Interi Sistema 
        $personalize = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        Set-ItemProperty -Path $personalize -Name "AppsUseLightTheme" -Value 0 -Type DWord
        Set-ItemProperty -Path $personalize -Name "SystemUsesLightTheme" -Value 0 -Type DWord

        # Abilitazione Output BSOD dettagliato utile lato debug
        $crashControl = "HKLM:\System\CurrentControlSet\Control\CrashControl"
        Set-ItemProperty -Path $crashControl -Name "DisplayParameters" -Value 1 -Type DWord

        $searchBox = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        Set-ItemProperty -Path $searchBox -Name "SearchboxTaskbarMode" -Value 0 -Type DWord

        $systemPol = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        Set-ItemProperty -Path $systemPol -Name "VerboseStatus" -Value 1 -Type DWord

        # Pulizia icone desktop base
        $desktopIcons = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
        if (-not (Test-Path $desktopIcons)) { New-Item -Path $desktopIcons -Force | Out-Null }
        Set-ItemProperty -Path $desktopIcons -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Value 0 -Type DWord
        Set-ItemProperty -Path $desktopIcons -Name "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}" -Value 0 -Type DWord
        Set-ItemProperty -Path $desktopIcons -Name "{645FF040-5081-101B-9F08-00AA002F954E}" -Value 0 -Type DWord
        Set-ItemProperty -Path $desktopIcons -Name "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}" -Value 0 -Type DWord
        
        Write-StyledMessage -Type Success -Text "✅ Personalizzazioni File Explorer applicate con successo."
    }
    catch {
        Write-StyledMessage -Type Warning -Text "⚠️ Errore durante l'applicazione delle chiavi di registro (Explorer): $($_.Exception.Message)"
    }
}

function Invoke-AdvancedTweaks {
    <#
    .SYNOPSIS
    Blocchi Registry duri contro telemetria, OneDrive, Edge e imposizione menu click destro classico.
    #>
    Write-StyledMessage -Type Info -Text "🛠️ Esecuzione tweak avanzati di sistema (Copilot, IPv4, Menu Classico)..."
    try {
        $cloudContent = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
        if (-not (Test-Path $cloudContent)) { New-Item -Path $cloudContent -Force | Out-Null }
        Set-ItemProperty -Path $cloudContent -Name "DisableWindowsConsumerFeatures" -Value 1 -Type DWord

        $taskbarDev = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
        if (-not (Test-Path $taskbarDev)) { New-Item -Path $taskbarDev -Force | Out-Null }
        Set-ItemProperty -Path $taskbarDev -Name "TaskbarEndTask" -Value 1 -Type DWord

        # Ripristino Vecchio Modello "Click Destro" Menu
        New-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" -Force -Value "" | Out-Null

        # Priorita IPv4 sulle interfacce (utile x reti legacy locali)
        $tcpip6 = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters"
        Set-ItemProperty -Path $tcpip6 -Name "DisabledComponents" -Value 32 -Type DWord

        # Rimozione feature Edge invadenti
        $edgeUpdate = "HKLM:\SOFTWARE\Policies\Microsoft\EdgeUpdate"
        if (-not (Test-Path $edgeUpdate)) { New-Item -Path $edgeUpdate -Force | Out-Null }
        Set-ItemProperty -Path $edgeUpdate -Name "CreateDesktopShortcutDefault" -Value 0 -Type DWord
        
        $edgePol = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
        if (-not (Test-Path $edgePol)) { New-Item -Path $edgePol -Force | Out-Null }
        Set-ItemProperty -Path $edgePol -Name "PersonalizationReportingEnabled" -Value 0 -Type DWord
        Set-ItemProperty -Path $edgePol -Name "ShowRecommendationsEnabled" -Value 0 -Type DWord
        Set-ItemProperty -Path $edgePol -Name "HideFirstRunExperience" -Value 1 -Type DWord
        Set-ItemProperty -Path $edgePol -Name "UserFeedbackAllowed" -Value 0 -Type DWord
        Set-ItemProperty -Path $edgePol -Name "ConfigureDoNotTrack" -Value 1 -Type DWord

        # Disabilitazione di Windows Copilot a basso livello
        $winCopilot = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"
        if (-not (Test-Path $winCopilot)) { New-Item -Path $winCopilot -Force | Out-Null }
        Set-ItemProperty -Path $winCopilot -Name "TurnOffWindowsCopilot" -Value 1 -Type DWord
        
        $winCopilotCU = "HKCU:\Software\Policies\Microsoft\Windows\WindowsCopilot"
        if (-not (Test-Path $winCopilotCU)) { New-Item -Path $winCopilotCU -Force | Out-Null }
        Set-ItemProperty -Path $winCopilotCU -Name "TurnOffWindowsCopilot" -Value 1 -Type DWord

        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowCopilotButton" -Value 0 -Type DWord

        # Estirpazione pulita del demone OneDrive
        Write-StyledMessage -Type Info -Text "🗑️ Disinstallazione profonda di Microsoft OneDrive in corso..."
        try {
            icacls $Env:OneDrive /deny "Administrators:(D,DC)" | Out-Null
            Start-Process 'C:\Windows\System32\OneDriveSetup.exe' -ArgumentList '/uninstall' -Wait -ErrorAction SilentlyContinue
            Stop-Process -Name FileCoAuth,Explorer -Force -ErrorAction SilentlyContinue
            Remove-Item "$Env:LocalAppData\Microsoft\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item "C:\ProgramData\Microsoft OneDrive" -Recurse -Force -ErrorAction SilentlyContinue
            icacls $Env:OneDrive /grant "Administrators:(D,DC)" | Out-Null
            Set-Service -Name OneSyncSvc -StartupType Disabled -ErrorAction SilentlyContinue
        } catch { }

        Write-StyledMessage -Type Success -Text "✅ Tweak avanzati di sistema applicati con successo."
    }
    catch {
        Write-StyledMessage -Type Warning -Text "⚠️ Errore durante i tweak avanzati: $($_.Exception.Message)"
    }
}

function Install-RequiredApps {
    <#
    .SYNOPSIS
    Installa il pacchetto applicativo selezionato da Roadmap silenziosamente tramite Winget in blocco sequenziale.
    #>
    $apps = @(
        @{ Id = "MartiCliment.UniGetUI"; Name = "Winget UI" },
        @{ Id = "Microsoft.PowerToys"; Name = "PowerToys" },
        @{ Id = "voidtools.Everything"; Name = "Everything" },
        @{ Id = "9N2DRHJ970D9"; Name = "EverythingPowerToys" },
        @{ Id = "nilesoft.shell"; Name = "Nilesoft Shell" }
    )

    foreach ($app in $apps) {
        Write-StyledMessage -Type Info -Text "⬇️ Installazione $($app.Name) tramite Winget in corso..."
        try {
            $process = Start-Process -FilePath winget -ArgumentList "install --id $($app.Id) --silent --accept-package-agreements --accept-source-agreements" -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -eq 0 -or $process.ExitCode -eq -1978335215) {
                Write-StyledMessage -Type Success -Text "✅ $($app.Name) installato o già presente nel sistema."
            } else {
                Write-StyledMessage -Type Warning -Text "⚠️ Installazione di $($app.Name) fallita (ExitCode: $($process.ExitCode))."
            }
        } catch {
            Write-StyledMessage -Type Warning -Text "⚠️ Errore d'installazione per $($app.Name): $($_.Exception.Message)"
        }
    }
}

function Deploy-CustomAssets {
    <#
    .SYNOPSIS
    Applica i temi custom di config alle app terze precedentemente installate scaricandoli dal repository.
    #>
    Write-StyledMessage -Type Info -Text "📦 Avvio download e applicazione asset pre-configurati..."
    
    $tempDir = $script:AppConfig.Paths.Temp
    if (-not (Test-Path $tempDir)) { New-Item -Path $tempDir -ItemType Directory -Force | Out-Null }

    # PowerToys Inject
    try {
        Write-StyledMessage -Type Info -Text "⚙️ Distribuzione configurazioni personalizzate PowerToys..."
        $ptZip = Join-Path $tempDir "PowerToys.zip"
        $ptExtract = Join-Path $tempDir "PowerToys_Extracted"
        Invoke-WebRequest -Uri $script:AppConfig.URLs.PowerToysConfig -OutFile $ptZip -UseBasicParsing
        Expand-Archive -Path $ptZip -DestinationPath $ptExtract -Force
        
        $ptDest = $script:AppConfig.Paths.PowerToysDir
        if (-not (Test-Path $ptDest)) { New-Item -Path $ptDest -ItemType Directory -Force | Out-Null }
        Copy-Item -Path "$ptExtract\*" -Destination $ptDest -Recurse -Force
        Write-StyledMessage -Type Success -Text "✅ Configurazioni PowerToys iniettate con successo."
    } catch {
        Write-StyledMessage -Type Warning -Text "⚠️ Impossibile configurare PowerToys: $($_.Exception.Message)"
    }

    # Nilesoft Shell Menu
    try {
        Write-StyledMessage -Type Info -Text "⚙️ Distribuzione configurazioni personalizzate Nilesoft Shell..."
        $nsZip = Join-Path $tempDir "NilesoftShell.zip"
        $nsExtract = Join-Path $tempDir "NilesoftShell_Extracted"
        Invoke-WebRequest -Uri $script:AppConfig.URLs.NilesoftConfig -OutFile $nsZip -UseBasicParsing
        Expand-Archive -Path $nsZip -DestinationPath $nsExtract -Force
        
        $nsDest = $script:AppConfig.Paths.NilesoftDir
        if (-not (Test-Path $nsDest)) { New-Item -Path $nsDest -ItemType Directory -Force | Out-Null }
        Copy-Item -Path "$nsExtract\*" -Destination $nsDest -Recurse -Force
        Write-StyledMessage -Type Success -Text "✅ Preset Nilesoft Shell applicato con successo."
    } catch {
        Write-StyledMessage -Type Warning -Text "⚠️ Impossibile configurare Nilesoft Shell: $($_.Exception.Message)"
    }
    
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

function Create-WinSupportShortcut {
    <#
    .SYNOPSIS
    Realizza l'icona ed il collegamento per il Supporto Remoto via RustDesk per l'utente finale.
    #>
    Write-StyledMessage -Type Info -Text "🔗 Creazione scorciatoia desktop per l'assistenza remota (Win Support)..."
    
    try {
        $desktop = $script:AppConfig.Paths.Desktop
        $shortcut = Join-Path $desktop "Win Support.lnk"
        $iconDir = $script:AppConfig.Paths.WinToolkitDir
        $icon = Join-Path $iconDir "WinSupport.png"
        $iconIco = Join-Path $iconDir "WinSupport.ico"

        if (-not (Test-Path $iconDir)) { New-Item -Path $iconDir -ItemType Directory -Force | Out-Null }

        if (-not (Test-Path $icon)) {
            Invoke-WebRequest -Uri $script:AppConfig.URLs.WinSupportIcon -OutFile $icon -UseBasicParsing
            Copy-Item $icon -Destination $iconIco -Force
        }

        # Genera Shortcut programmatica al file wt.exe eseguendo il payload
        $shell = New-Object -ComObject WScript.Shell
        $link = $shell.CreateShortcut($shortcut)
        $link.TargetPath = $script:AppConfig.Paths.wtExe
        $link.Arguments = 'pwsh -NoProfile -ExecutionPolicy Bypass -Command "irm installazione RustDesk | iex"'
        $link.WorkingDirectory = $script:AppConfig.Paths.WinToolkitDir
        if (Test-Path $iconIco) { $link.IconLocation = "$iconIco,0" }
        $link.Description = "Assistenza Win Support"
        $link.Save()

        Write-StyledMessage -Type Success -Text "✅ Collegamento 'Win Support' generato sul Desktop."
    }
    catch {
        Write-StyledMessage -Type Error -Text "❌ Errore durante la creazione del collegamento Desktop: $($_.Exception.Message)"
    }
}


# ============================================================================
# AVVIO PRINCIPALE E GESTIONE TRANSIZIONI
# ============================================================================

function Invoke-WinStarterSetup {
    <#
    .SYNOPSIS
    Inietta la logica dell'intero script in sequenza. Esegue le verifiche di policy, pre-requisti, log
    e lancia man mano tutte le sotto-funzioni in ordine, gestendo alla perfezione anche lo switch a PS7+.
    #>
    try {
        $isResumeSetup = $env:WINTOOLKIT_RESUME -eq "1"
        $Host.UI.RawUI.WindowTitle = "Win Starter by MagnetarMan"

        Start-ToolkitLog "WinStarter"

        # Check UAC for Administrator Rights e Re-Esecuzione nativa se necessario
        if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-StyledMessage -Type Warning -Text "⚠️ Privilegi insufficienti. Riavvio dello script come Amministratore..."
            $procParams = @{
                FilePath     = 'powershell'
                ArgumentList = @( '-ExecutionPolicy', 'Bypass', '-NoProfile', '-File', "`"$PSCommandPath`"" )
                Verb         = 'RunAs'
            }
            Start-Process @procParams
            exit
        }

        Show-Header -Title $script:AppConfig.Header.Title -Version $script:AppConfig.Header.Version

        # Logica Primaria Pre-Transizione
        if (-not $isResumeSetup) {
            Write-StyledMessage -Type Info -Text "✨ Avvio inizializzazione ambiente Win Starter..."
            
            Update-EnvironmentPath
            if (-not (Test-WingetFunctionality)) {
                Write-StyledMessage -Type Warning -Text "⚠️ Winget non risponde. Auto-riparazione..."
                Install-WingetCore
                Repair-WingetDatabase
                Update-EnvironmentPath
            }
            Test-WingetDeepValidation | Out-Null
            Install-PowerShellCore | Out-Null
        }

        # Transizione a PowerShell 7 se avviato in 5.1 e la 7 è disponibile
        if ($PSVersionTable.PSVersion.Major -lt 7 -and (Test-Path "$env:ProgramFiles\PowerShell\7\pwsh.exe")) {
            Write-StyledMessage -Type Info -Text "✨ Rilevata PowerShell 7. Upgrade dell'ambiente di esecuzione in progress..."
            $env:WINTOOLKIT_RESUME = "1"
            $ps7Path = $script:AppConfig.Paths.PowerShell7
            $procParams = @{
                FilePath     = Join-Path $ps7Path "pwsh.exe"
                ArgumentList = @("-ExecutionPolicy", "Bypass", "-NoExit", "-File", "`"$PSCommandPath`"")
                Verb         = "RunAs"
            }
            Start-Process @procParams
            exit
        }

        # Setup Essenziale Windows (Terminale e Ambiente Custom)
        $wtInstalled = Install-WindowsTerminalApp
        if ($wtInstalled -and (Get-Command 'wt.exe' -ErrorAction SilentlyContinue)) {
            Write-StyledMessage -Type Info -Text "⚙️ Impostazione Windows Terminal come terminale predefinito..."
            try {
                $registryPath = $script:AppConfig.Registry.TerminalStartup
                if (-not (Test-Path $registryPath)) { $null = New-Item -Path $registryPath -Force }
                Set-ItemProperty -Path $registryPath -Name 'DelegationTerminal' -Value '{E12F0936-0E6F-548E-A9F6-B20C69A27D17}' -Force
                Set-ItemProperty -Path $registryPath -Name 'DelegationConsole' -Value '{B23D10C0-31E3-401A-97EF-4BB30B62E10B}' -Force
            } catch { }
        }
        Install-PspEnvironment

        # Disinnesco Modifiche Sgradite (OS Baseline)
        SetRecommendedUpdate
        Set-ExplorerPersonalization
        Invoke-AdvancedTweaks
        
        Write-StyledMessage -Type Info -Text "🔄 Riavvio processo Explorer.exe per rendere effettive le modifiche UI..."
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Sleep 2

        # Fase di Deploy Pacchetti e Asset Winstarter
        Install-RequiredApps
        Deploy-CustomAssets
        Create-WinSupportShortcut

        # Se siamo in terminal Host ma Windows Terminal è stato installato sposta l'utente la dentro alla fine
        if (-not ($env:WT_SESSION) -and (Get-Command "wt.exe" -ErrorAction SilentlyContinue)) {
            Write-StyledMessage -Type Info -Text "Riavvio finale in Windows Terminal..."
            try {
                $wtArgs = "-w 0 new-tab -p `"PowerShell`" -d . pwsh.exe -ExecutionPolicy Bypass -NoExit -Command `"Write-Host '✅ Ambiente Pronto. Configurazione Winstarter conclusa!' -ForegroundColor Green`""
                Start-Process -FilePath "wt.exe" -ArgumentList $wtArgs
                exit
            } catch { }
        }

        Write-StyledMessage -Type Success -Text "🎉 Configurazione Win Starter conclusa brillantemente! Il sistema è pronto."
        Write-Host "Premi un tasto per uscire..."
        $null = [Console]::ReadKey($true)
        exit 0
    }
    catch {
        Write-StyledMessage -Type Error -Text "❌ Errore critico irreversibile nel setup: $($_.Exception.Message)"
        Write-ToolkitLog -Level 'ERROR' -Message "ECCEZIONE UNHANDLED: $($_.Exception.Message) `n $($_.ScriptStackTrace)"
        Write-Host "Premi un tasto per uscire..."
        $null = [Console]::ReadKey($true)
        exit 1
    }
}

Invoke-WinStarterSetup
