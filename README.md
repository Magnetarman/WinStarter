<p align="center">
<img src="/img/Win-Starter.png" alt="Win-Starter-Icon" width="180">
</p>

# Win Starter: Configura Windows con un solo comando

## _The MagnetarMan Way_

<p>
	<img src="https://img.shields.io/github/license/Magnetarman/WinStarter?style=for-the-badge&logo=opensourceinitiative&logoColor=white&color=0080ff" alt="license">
	<img src="https://img.shields.io/badge/version-1.3.1-green.svg?style=for-the-badge" alt="versione">
	<img src="https://img.shields.io/github/last-commit/Magnetarman/WinStarter?style=for-the-badge&logo=git&logoColor=white&color=9370DB" alt="last-commit">
</p>

Win Starter è uno script PowerShell che automatizza la configurazione iniziale di Windows: verifica e ripara **Winget**, installa software essenziali (PowerToys, Everything, Nilesoft Shell, UniGet UI), applica **tweak di sistema** per ridurre bloatware e telemetria, e prepara un ambiente terminale moderno con PowerShell 7, Windows Terminal, Oh My Posh e strumenti da riga di comando. Un solo comando per portare il sistema allo stato desiderato.

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

6. Segui le istruzioni a video; al termine avrai un ambiente Windows ottimizzato e già configurato.

## 🐛 Bug Noti - Fix in corso…

- Barre di progressione non correttamente soppresse nell'output
- PowerToys avvia lo splash screen generale, dovrebbe avviarsi ridotto ad icona senza splash screen
- Disabilitazione delle notifiche di PowerToys nelle notifiche di windows

---

## 👾 Componenti e funzioni

Lo script è organizzato in blocchi funzionali. Di seguito una guida al flusso di esecuzione e alle funzioni principali.

### Sicurezza e Windows Update

| Funzione                                   | Descrizione                                                                                                                                           |
| ------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Blocco Connessione a Consumo**           | Identifica la rete attiva e la imposta temporaneamente come "A Consumo" (`Cost=2`) per bloccare il download di aggiornamenti durante l'esecuzione.    |
| **Impostazioni Windows Update**            | Blocca aggiornamenti driver, posticipa Feature/Quality updates (365/4 gg) e disabilita il riavvio automatico via Registry.                              |
| **Riavvio Servizio WU**                    | Riavvia forzatamente `wuauserv` all'inizio per applicare immediatamente le nuove policy di aggiornamento.                                             |
| **Ripristino Connessione**                 | Al termine del setup, ripristina la rete su "Illimitata" (`Cost=1`) per consentire il normale funzionamento.                                          |

### Winget e Riparazione

| Funzione                       | Descrizione                                                                                             |
| ------------------------------ | ------------------------------------------------------------------------------------------------------- |
| **Test-WingetFunctionality**   | Verifica che Winget risponda correttamente nel PATH.                                                    |
| **Repair-WingetDatabase**      | Ripristina il database Winget se corrotto: pulisce cache, resetta sorgenti e corregge permessi.         |
| **Test-WingetDeepValidation**  | Esegue un test di ricerca reale e, in caso di crash (`ACCESS_VIOLATION`), avvia la riparazione drastica. |
| **Install-WingetCore**         | Fallback finale: reinstalla Winget scaricando il bundle MSIX ufficiale da GitHub.                       |

### Ambiente Shell e App

| Funzione                       | Descrizione                                                                                                                                                |
| ------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Install-PowerShellCore**     | Installa PowerShell 7 e migra l'esecuzione dal PowerShell 5.1 legacy per supportare i moduli moderni.                                                      |
| **Install-WindowsTerminalApp** | Installa l'app Windows Terminal e la imposta come terminale predefinito di sistema.                                                                        |
| **Install-PspEnvironment**     | Configura Oh My Posh, zoxide, btop, fastfetch e i font Nerd; scarica profilo e temi personalizzati dal repository.                                          |
| **Install-RequiredApps**       | Installa in blocco: **UniGet UI**, **PowerToys**, **Everything** e **Nilesoft Shell**, pulendo eventuali shortcut superflue dal desktop.                   |
| **Deploy-CustomAssets**        | Applica preset e configurazioni per PowerToys e Nilesoft Shell iniettando i file nelle directory di sistema.                                               |

### Tweak e Personalizzazione

| Funzione                        | Descrizione                                                                                                                                              |
| ------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Set-ExplorerPersonalization** | Abilita estensioni file, Dark Mode, visualizzazione BSOD dettagliata e pulisce icone desktop predefinite.                                                |
| **Invoke-AdvancedTweaks**       | Tweak profondi: disabilita Copilot, OneDrive, telemetria Edge, rimuove Teams Free e ripristina il **Menu Contestuale Classico** (Windows 10 style).      |
| **Restart-ExplorerSafe**        | Riavvia la shell `explorer.exe` in modo sicuro per rendere effettive tutte le modifiche visive e al registro.                                            |

---

## 📁 Percorsi importanti

| Descrizione                | Percorso                             |
| -------------------------- | ------------------------------------ |
| Log di Win Starter         | `%LOCALAPPDATA%\WinStarter\logs`     |
| Directory temporanea setup | `%TEMP%\WinStarterSetup`             |
| Configurazione PowerToys   | `%LOCALAPPDATA%\Microsoft\PowerToys` |
| Nilesoft Shell             | `%ProgramFiles%\Nilesoft Shell`      |

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
