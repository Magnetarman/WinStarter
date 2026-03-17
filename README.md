<p align="center">
<img src="/img/Win-Starter.png" alt="Win-Starter-Icon" width="180">
</p>

# Win Starter: Configura Windows con un solo comando

## _The MagnetarMan Way_

<p>
	<img src="https://img.shields.io/github/license/Magnetarman/WinStarter?style=for-the-badge&logo=opensourceinitiative&logoColor=white&color=0080ff" alt="license">
	<img src="https://img.shields.io/badge/version-1.2.3-green.svg?style=for-the-badge" alt="versione">
	<img src="https://img.shields.io/github/last-commit/Magnetarman/WinStarter?style=for-the-badge&logo=git&logoColor=white&color=9370DB" alt="last-commit">
</p>

Win Starter è uno script PowerShell che automatizza la configurazione iniziale di Windows: verifica e ripara **Winget**, installa software essenziali (PowerToys, Everything, Nilesoft Shell, UniGet UI), applica **tweak di sistema** per ridurre bloatware e telemetria, e prepara un ambiente terminale moderno con PowerShell 7, Windows Terminal, Oh My Posh e strumenti da riga di comando. Infine inserisce l'icona di WinToolkit sul desktop e l'icona per effettuare il deploy di Rust Desk con una configurazione personalizzata per effettuare supporto remoto. Un solo comando per portare il sistema allo stato desiderato.

---

## ⚙️ Requisiti minimi

> [!IMPORTANT]
>
> Prima di avviare Win Starter, assicurati di soddisfare i seguenti requisiti:
>
> - **Connessione a Internet** (per download di Winget, app e asset);
> - **Esecuzione come Amministratore** (lo script richiederà il riavvio con privilegi elevati se necessario);
> - **Disattiva temporaneamente Windows Defender (Protezione in tempo reale)**: alcuni passaggi (soprattutto la riparazione di Winget) possono essere bloccati da **falsi positivi** e portare a fallimenti “catastrofici” della riparazione. Al termine **riattivalo**.
> - **Windows 10** (build 16299 o superiore) oppure **Windows 11**.
>
> Inoltre, per evitare blocchi o falsi positivi (soprattutto durante la riparazione di Winget e l’installazione di pacchetti MSIX/AppX), è **consigliato disattivare temporaneamente Microsoft Defender / SmartScreen** prima dell’esecuzione e riattivarlo a fine operazione.

| Versioni di Windows | Supportato                    |
| :------------------ | :---------------------------- |
| Windows 11          | 🟢 Sì                         |
| Windows 10 >= 1709  | 🟢 Sì                         |
| Windows 10 < 1709   | 🔴 No (Winget non supportato) |

---

## 🚀 Come eseguire Win Starter

1. Premi il tasto **Windows** sulla tastiera oppure apri la ricerca di Windows.
2. Digita **PowerShell** nel campo di ricerca.
3. Clicca con il tasto destro su **Windows PowerShell**.
4. Seleziona **Esegui come amministratore**.
5. Copia e incolla nella finestra di PowerShell il comando seguente:

```powershell
irm https://magnetarman.com/winstarter | iex
```

6. Segui le istruzioni a video; al termine troverai sul Desktop la scorciatoia **Win Support** (assistenza remota) e un ambiente già configurato.

## 🪱 Bug Noti - Fix in corso…

- Barre di progressione non correttamente soppresse nell'output
- PowerToys avvia lo splash screen generale, dovrebbe avviarsi ridotto ad icona senza splash screen
- Disabilitazione delle notifiche di PowerToys
- 

---

## 👾 Componenti e funzioni

Lo script è organizzato in blocchi funzionali. Di seguito una guida alle funzioni principali e al flusso di esecuzione.

### Utilità base e presentazione

| Funzione                                    | Descrizione                                                            |
| ------------------------------------------- | ---------------------------------------------------------------------- |
| **Format-CenteredText**                     | Centra il testo nel terminale per l’intestazione.                      |
| **Show-Header**                             | Pulisce lo schermo e mostra il banner ASCII "Win Starter".             |
| **Write-StyledMessage**                     | Stampa messaggi con icone (✅ ⚠️ ❌ 💎) e li registra nel file di log. |
| **Start-ToolkitLog** / **Write-ToolkitLog** | Inizializza e scrive il log in `%LOCALAPPDATA%\WinStarter\logs`.       |

### Winget e permessi

| Funzione                                                | Descrizione                                                                                             |
| ------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| **Start-AppxSilentProcess**                             | Installa pacchetti AppX/MSIX in background senza finestre di progresso.                                 |
| **Stop-InterferingProcess**                             | Termina processi che bloccano Winget (Store, AppInstaller, winget, ecc.).                               |
| **Invoke-ForceCloseWinget**                             | Chiude tutti i processi Winget per liberare lock sui file.                                              |
| **Update-EnvironmentPath**                              | Ricarica il PATH nel processo corrente dopo nuove installazioni.                                        |
| **Set-PathPermissions** / **Set-WingetPathPermissions** | Corregge i permessi sulla cartella di installazione di Winget (Administrators = FullControl).           |
| **Invoke-WingetCommand**                                | Esegue comandi Winget con timeout e flag `--disable-interactivity` se supportato.                       |
| **Test-WingetFunctionality**                            | Verifica che Winget sia nel PATH e risponda (es. `winget --version`).                                   |
| **Test-WingetCompatibility**                            | Controlla che la build di Windows supporti Winget (>= 16299).                                           |
| **Repair-WingetDatabase**                               | Ripristina un database Winget corrotto: pulisce cache, resetta sorgenti e permessi.                     |
| **Test-WingetDeepValidation**                           | Simula un uso reale (es. `winget search`) e, in caso di crash (ACCESS_VIOLATION), avvia la riparazione. |
| **Install-WingetCore**                                  | Rimedio estremo: scarica il bundle MSIX da GitHub e reinstalla Winget.                                  |

### Installazione componenti sistema

| Funzione                       | Descrizione                                                                                                                                                                                                                                   |
| ------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Install-PowerShellCore**     | Installa PowerShell 7 tramite Winget se non presente; necessario per evitare limiti della 5.1.                                                                                                                                                |
| **Install-WindowsTerminalApp** | Installa Windows Terminal (ID 9N0DX20HK701) da Microsoft Store tramite Winget.                                                                                                                                                                |
| **Install-PspEnvironment**     | Configura l’ambiente shell: installa **Oh My Posh**, **zoxide**, **btop**, **fastfetch**, **JetBrains Mono Nerd Font**; scarica tema Oh My Posh (atomic), profilo PowerShell e `settings.json` di Windows Terminal dal repository WinToolkit. |

### Funzioni core Win Starter

| Funzione                        | Descrizione                                                                                                                                                                                                                                                                                                                                         |
| ------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **SetRecommendedUpdate**        | Mitiga Windows Update: disabilita aggiornamenti driver da WU, posticipa Feature Update (365 gg) e Quality Update (4 gg), disabilita riavvio automatico e metadati driver da rete.                                                                                                                                                                   |
| **Set-ExplorerPersonalization** | Personalizza Esplora file e shell: estensioni visibili, cartella predefinita “Questo PC”, dark mode (app e sistema), rimozione suggerimenti Bing dalla ricerca, BSOD dettagliato, icone desktop (Computer, Rete, Cestino, ecc.).                                                                                                                    |
| **Invoke-AdvancedTweaks**       | Tweak avanzati: disabilita funzionalità consumer (CloudContent), “Chiudi” dalla barra delle applicazioni, **menu contestuale classico** (click destro), priorità IPv4, limiti a Edge (no shortcut desktop, no personalization/reporting, Do Not Track), **disabilitazione Windows Copilot** (HKLM + HKCU), **disinstallazione e pulizia OneDrive**. |
| **Install-RequiredApps**        | Installa in blocco tramite Winget: **UniGet UI**, **PowerToys**, **Everything**, **Everything PowerToys**, **Nilesoft Shell**.                                                                                                                                                                                                                      |
| **Deploy-CustomAssets**         | Scarica dal repository gli asset preconfigurati (PowerToys.zip, NilesoftShell.zip), li estrae e li copia nelle cartelle di PowerToys e Nilesoft Shell.                                                                                                                                                                                              |
| **Create-WinSupportShortcut**   | Crea sul Desktop il collegamento **Win Support** che avvia Windows Terminal con comando per l’installazione/assistenza remota (RustDesk); usa l’icona scaricata da repository.                                                                                                                                                                      |

### Flusso principale: Invoke-WinStarterSetup

1. **Controllo privilegi**: se non si è amministratori, lo script si riavvia con `RunAs`.
2. **Pre-transizione**: aggiorna PATH, verifica Winget; se non funziona, installa/ripara Winget e esegue la validazione “deep”.
3. **Transizione a PowerShell 7**: se lo script è partito con PowerShell 5.1 e PowerShell 7 è installato, si riavvia in `pwsh.exe` (con `WINTOOLKIT_RESUME=1`) per proseguire in ambiente moderno.
4. **Setup essenziale**: installa Windows Terminal e lo imposta come terminale predefinito (registro), poi esegue **Install-PspEnvironment**.
5. **Baseline OS**: applica **SetRecommendedUpdate**, **Set-ExplorerPersonalization**, **Invoke-AdvancedTweaks**; riavvia Explorer.
6. **Deploy**: **Install-RequiredApps**, **Deploy-CustomAssets**, **Create-WinSupportShortcut**.
7. **Chiusura**: se non si è già in Windows Terminal, apre una nuova finestra WT con messaggio di completamento; altrimenti mostra messaggio e attende un tasto.

---

## 📁 Percorsi importanti

| Descrizione                  | Percorso                                            |
| ---------------------------- | --------------------------------------------------- |
| Log di Win Starter           | `%LOCALAPPDATA%\WinStarter\logs`                    |
| Directory temporanea setup   | `%TEMP%\WinStarterSetup`                            |
| Configurazione PowerToys     | `%LOCALAPPDATA%\Microsoft\PowerToys`                |
| Nilesoft Shell               | `%ProgramFiles%\Nilesoft Shell`                     |
| Icona / contesto Win Support | `%LOCALAPPDATA%\WinToolkit` (WinSupport.png / .ico) |

---

## 🤔 Domande frequenti

### A cosa serve Win Starter?

A portare un’installazione Windows “pulita” allo stato desiderato con un solo comando: Winget funzionante, PowerShell 7 e terminale moderno, app essenziali installate, configurazioni PowerToys e Nilesoft applicate, riduzione bloat (OneDrive, Copilot, aggiornamenti driver invasivi) e ambiente shell curato (Oh My Posh, zoxide, btop, fastfetch, font Nerd).

### Perché a volte lo script si riavvia?

- **Primo riavvio**: per acquisire diritti di Amministratore se avviato senza.
- **Secondo riavvio**: per passare da PowerShell 5.1 a PowerShell 7 quando disponibile, così tutto il resto gira in ambiente moderno.

### Posso eseguire solo alcune parti?

Lo script è pensato per un’esecuzione end-to-end. Per usare singole funzioni apri `winstarter.ps1` in PowerShell (come Amministratore), dot-source il file (`. .\winstarter.ps1`) e invoca le funzioni che ti servono.

### Dove sono i file di log?

In `%LOCALAPPDATA%\WinStarter\logs`, con nome tipo `WinStarter_yyyy-MM-dd_HH-mm-ss.log`.

---

## 🎗 Autore

Creato con ❤️ da [Magnetarman](https://magnetarman.com/).
