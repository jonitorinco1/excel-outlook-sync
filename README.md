# Excel → Outlook Calendar Sync

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)
![Outlook](https://img.shields.io/badge/Microsoft-Outlook-0072C6?logo=microsoft-outlook)

Sincronizzazione automatica di scadenze da un file **Excel** verso il **Calendario di Microsoft Outlook** tramite interfaccia COM (`pywin32`).

---

## Indice

- [Funzionalità](#funzionalità)
- [Requisiti](#requisiti)
- [Struttura del progetto](#struttura-del-progetto)
- [Installazione](#installazione)
- [Configurazione](#configurazione)
- [Struttura del file Excel](#struttura-del-file-excel)
- [Utilizzo](#utilizzo)
- [Automazione con Task Scheduler](#automazione-con-task-scheduler)
- [Log](#log)
- [Logica di sincronizzazione](#logica-di-sincronizzazione)
- [Licenza](#licenza)

---

## Funzionalità

- Legge scadenze da un file Excel (colonne configurabili)
- Crea automaticamente eventi nel calendario Outlook se non esistono
- Aggiorna gli eventi se i dati sono cambiati (titolo, data, descrizione, categoria)
- **Non elimina mai** eventi già presenti in Outlook (comportamento conservativo)
- Identifica gli eventi tramite un tag univoco `[REF:SC-001]` nel corpo dell'appuntamento
- Produce un file di log giornaliero con il riepilogo delle operazioni
- Schedulabile con **Task Scheduler di Windows**

---

## Requisiti

| Requisito | Versione minima |
|-----------|----------------|
| Windows | 10 / 11 |
| Python | 3.9+ |
| Microsoft Outlook | Desktop (qualsiasi versione moderna) |
| `pandas` | 2.0.0+ |
| `openpyxl` | 3.1.0+ |
| `pywin32` | 306+ |

> ⚠️ Lo script utilizza l'interfaccia COM di Windows. **Non funziona** con Outlook Web, Microsoft 365 web-only o su macOS/Linux.

---

## Struttura del progetto

```
excel-outlook-sync/
├── sync.py                    # Script principale di sincronizzazione
├── config.ini.example         # Configurazione di esempio (da copiare in config.ini)
├── requirements.txt           # Dipendenze Python
├── run_sync.bat               # Batch per avvio manuale e Task Scheduler
├── README.md                  # Questo file
├── .gitignore                 # File esclusi da git
├── logs/                      # Log giornalieri (generati a runtime)
│   └── .gitkeep
└── template/
    └── scadenze_template.xlsx # Template Excel di esempio
```

---

## Installazione

### 1. Clona il repository

```bash
git clone https://github.com/jonitorinco1/excel-outlook-sync.git
cd excel-outlook-sync
```

### 2. (Consigliato) Crea un ambiente virtuale

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Installa le dipendenze

```bash
pip install -r requirements.txt
```

### 4. Post-installazione pywin32

Dopo aver installato `pywin32`, esegui questo comando **una sola volta** (richiede privilegi di amministratore):

```bash
python -m pywin32_postinstall -install
```

---

## Configurazione

### 1. Crea il file di configurazione

```bash
copy config.ini.example config.ini
```

### 2. Modifica `config.ini`

```ini
[Excel]
percorso_file = C:\Users\TuoNome\Documents\scadenze.xlsx
nome_foglio = Scadenze

[Outlook]
nome_calendario = Calendario
durata_default_minuti = 60

[Log]
cartella_log = logs
livello = INFO
```

| Parametro | Descrizione |
|-----------|-------------|
| `percorso_file` | Percorso completo al file Excel |
| `nome_foglio` | Nome del foglio Excel (default: `Scadenze`) |
| `nome_calendario` | Nome del calendario in Outlook (es. `Calendario`, `Calendar`) |
| `durata_default_minuti` | Durata appuntamenti in minuti (default: `60`) |
| `livello` | Verbosità log: `DEBUG`, `INFO`, `WARNING`, `ERROR` |

> `config.ini` è escluso da `.gitignore` per proteggere i percorsi personali.

---

## Struttura del file Excel

Il file Excel deve contenere un foglio con le seguenti colonne (la riga 1 deve essere l'intestazione):

| Colonna | Obbligatoria | Descrizione | Esempio |
|---------|:---:|-------------|---------|
| `Riferimento` | ✅ | ID univoco della scadenza | `SC-001` |
| `Titolo` | ✅ | Oggetto dell'appuntamento | `Rinnovo contratto fornitore` |
| `Data` | ✅ | Data della scadenza | `31/03/2025` |
| `Ora` | ❌ | Ora dell'appuntamento | `10:00` |
| `Descrizione` | ❌ | Testo del corpo dell'appuntamento | `Contattare il fornitore XYZ` |
| `Categoria` | ❌ | Categoria colore in Outlook | `Contratti` |

Un template di esempio è disponibile in [`template/scadenze_template.xlsx`](template/scadenze_template.xlsx).

> Il **Riferimento** deve essere univoco per ogni riga: viene usato come chiave di sincronizzazione tra Excel e Outlook tramite il tag `[REF:SC-001]` inserito nel corpo dell'appuntamento.

---

## Utilizzo

### Esecuzione manuale (da terminale)

```bash
# Con ambiente virtuale attivo
python sync.py
```

### Esecuzione tramite file batch

```bash
run_sync.bat manual
```

Il parametro `manual` mostra un prompt al termine — utile per il debug. Omettilo per l'automazione.

---

## Automazione con Task Scheduler

### Configurazione via interfaccia grafica

1. Apri **Task Scheduler** (cerca "Utilità di pianificazione" nel menu Start)
2. Clic su **Crea attività di base...**
3. **Nome**: `Sync Excel Outlook`
4. **Trigger**: Giornaliero, all'ora desiderata (es. 08:00)
5. **Azione**: Avvia programma
   - Programma: percorso completo a `run_sync.bat`
     es. `C:\Users\TuoNome\project\excel-outlook-sync\run_sync.bat`
   - Argomenti: *(lascia vuoto)*
   - Inizia in: `C:\Users\TuoNome\project\excel-outlook-sync`
6. ✅ Spunta **Esegui che l'utente sia connesso o meno** (opzionale)
7. Clic **Fine**

### Configurazione via PowerShell (alternativa)

Esegui PowerShell come **Amministratore**:

```powershell
$action = New-ScheduledTaskAction `
    -Execute "C:\Users\TuoNome\project\excel-outlook-sync\run_sync.bat"

$trigger = New-ScheduledTaskTrigger -Daily -At 08:00

Register-ScheduledTask `
    -TaskName "Sync Excel Outlook" `
    -Action $action `
    -Trigger $trigger `
    -RunLevel Highest `
    -Description "Sincronizza scadenze Excel con Outlook"
```

> ⚠️ Outlook deve essere aperto (o avviabile in background) al momento dell'esecuzione.
> Per la modalità "non connesso", assicurati che le credenziali siano memorizzate.

---

## Log

I log vengono salvati nella cartella `logs/` con un file per ogni giorno di esecuzione:

```
logs/
├── sync_20250331.log
├── sync_20250401.log
└── ...
```

**Formato del log:**

```
2025-03-31 08:00:01 [INFO] ============================================================
2025-03-31 08:00:01 [INFO] AVVIO SINCRONIZZAZIONE Excel → Outlook
2025-03-31 08:00:01 [INFO] ============================================================
2025-03-31 08:00:02 [INFO] Lettura file Excel: C:\...\scadenze.xlsx (foglio: 'Scadenze')
2025-03-31 08:00:02 [INFO] Righe valide trovate nell'Excel: 12
2025-03-31 08:00:03 [INFO] Calendario trovato: 'Calendario'
2025-03-31 08:00:05 [INFO] [SC-001] CREATO: 'Rinnovo contratto' il 15/04/2025 09:00
2025-03-31 08:00:06 [INFO] [SC-002] AGGIORNATO: 'Revisione bilancio' il 30/04/2025 14:00
2025-03-31 08:00:06 [INFO] ------------------------------------------------------------
2025-03-31 08:00:06 [INFO] RIEPILOGO SINCRONIZZAZIONE:
2025-03-31 08:00:06 [INFO]   Creati    : 1
2025-03-31 08:00:06 [INFO]   Aggiornati: 1
2025-03-31 08:00:06 [INFO]   Invariati : 10
2025-03-31 08:00:06 [INFO]   Errori    : 0
2025-03-31 08:00:06 [INFO] ============================================================
2025-03-31 08:00:06 [INFO] Sincronizzazione completata.
```

Imposta `livello = DEBUG` in `config.ini` per visualizzare anche gli eventi invariati.

---

## Logica di sincronizzazione

```
Per ogni riga Excel:
  ├── Cerca in Outlook un evento con [REF:ID] nel corpo
  │
  ├── Non trovato → CREA nuovo evento
  │
  ├── Trovato + dati cambiati → AGGIORNA evento esistente
  │
  └── Trovato + dati identici → nessuna azione (INVARIATO)

Gli eventi non più presenti in Excel NON vengono eliminati da Outlook.
```

Il tag `[REF:SC-001]` viene inserito automaticamente in fondo al corpo di ogni appuntamento e usato come chiave di ricerca ad ogni sincronizzazione.

---

## Licenza

Distribuito sotto licenza **MIT**. Vedi [LICENSE](LICENSE) per i dettagli.

---

*Progetto creato per automatizzare la gestione delle scadenze aziendali su Microsoft Outlook.*
*Repository: [github.com/jonitorinco1/excel-outlook-sync](https://github.com/jonitorinco1/excel-outlook-sync)*
