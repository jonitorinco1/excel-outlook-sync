"""
sync.py - Sincronizzazione scadenze Excel → Calendario Outlook
==============================================================
Autore: [Il tuo nome]
Licenza: MIT
Descrizione:
    Legge un file Excel contenente scadenze e le sincronizza con un
    calendario di Microsoft Outlook tramite interfaccia COM (win32com).
    Supporta creazione e aggiornamento degli eventi; non elimina mai
    eventi già presenti in Outlook (comportamento conservativo).

Esecuzione:
    python sync.py
    oppure tramite run_sync.bat schedulato con Task Scheduler di Windows.
"""

import configparser
import logging
import os
import sys
from datetime import datetime, date, time as dt_time
from pathlib import Path

import pandas as pd

# ── Importazione win32com (richiede pywin32 installato + Outlook Desktop) ──
try:
    import win32com.client
    import pywintypes
except ImportError:
    print("[ERRORE] Il modulo 'pywin32' non è installato. Esegui: pip install pywin32")
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────────────────
# COSTANTI
# ─────────────────────────────────────────────────────────────────────────────
# Prefisso tag usato nel corpo dell'appuntamento per identificarlo univocamente
TAG_PREFIX = "[REF:"
TAG_SUFFIX = "]"

# Numero di giorni indietro/avanti entro cui cercare eventi esistenti
# (ottimizzazione: evita di scansionare l'intero calendario)
SEARCH_WINDOW_DAYS = 365

# Formato data/ora nei log
LOG_DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"

# ─────────────────────────────────────────────────────────────────────────────
# CARICAMENTO CONFIGURAZIONE
# ─────────────────────────────────────────────────────────────────────────────

def carica_configurazione(config_path: str = "config.ini") -> configparser.ConfigParser:
    """Carica il file di configurazione config.ini."""
    config = configparser.ConfigParser()

    if not os.path.exists(config_path):
        print(f"[ERRORE] File di configurazione '{config_path}' non trovato.")
        print("  → Copia 'config.ini.example' in 'config.ini' e personalizzalo.")
        sys.exit(1)

    config.read(config_path, encoding="utf-8")
    return config


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAZIONE LOGGING
# ─────────────────────────────────────────────────────────────────────────────

def configura_logging(log_dir: str, log_level: str = "INFO") -> logging.Logger:
    """
    Configura il logger per scrivere sia su file (nella cartella logs/)
    sia sulla console.
    """
    Path(log_dir).mkdir(parents=True, exist_ok=True)

    # Nome file di log con data odierna (un file al giorno)
    log_filename = datetime.now().strftime("sync_%Y%m%d.log")
    log_filepath = os.path.join(log_dir, log_filename)

    livello = getattr(logging, log_level.upper(), logging.INFO)

    # Formato messaggio di log
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt=LOG_DATETIME_FORMAT,
    )

    logger = logging.getLogger("excel_outlook_sync")
    logger.setLevel(livello)

    # Handler file
    file_handler = logging.FileHandler(log_filepath, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Handler console
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger


# ─────────────────────────────────────────────────────────────────────────────
# LETTURA EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def leggi_excel(percorso_file: str, nome_foglio: str, logger: logging.Logger) -> pd.DataFrame:
    """
    Legge il file Excel e restituisce un DataFrame pandas.
    Colonne attese (case-insensitive):
        Riferimento, Titolo, Data, Ora (opzionale),
        Descrizione (opzionale), Categoria (opzionale)
    """
    if not os.path.exists(percorso_file):
        logger.error(f"File Excel non trovato: {percorso_file}")
        sys.exit(1)

    logger.info(f"Lettura file Excel: {percorso_file} (foglio: '{nome_foglio}')")

    try:
        df = pd.read_excel(percorso_file, sheet_name=nome_foglio, dtype=str)
    except Exception as e:
        logger.error(f"Impossibile leggere il file Excel: {e}")
        sys.exit(1)

    # Normalizza i nomi delle colonne (rimuovi spazi, converti in Title Case)
    df.columns = [c.strip().title() for c in df.columns]

    # Verifica colonne obbligatorie
    colonne_obbligatorie = {"Riferimento", "Titolo", "Data"}
    colonne_mancanti = colonne_obbligatorie - set(df.columns)
    if colonne_mancanti:
        logger.error(f"Colonne obbligatorie mancanti nel foglio Excel: {colonne_mancanti}")
        sys.exit(1)

    # Aggiungi colonne opzionali vuote se mancanti
    for col in ("Ora", "Descrizione", "Categoria"):
        if col not in df.columns:
            df[col] = ""

    # Rimuovi righe senza Riferimento o Titolo
    df = df.dropna(subset=["Riferimento", "Titolo"])
    df = df[df["Riferimento"].str.strip() != ""]
    df = df[df["Titolo"].str.strip() != ""]

    logger.info(f"Righe valide trovate nell'Excel: {len(df)}")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# UTILITÀ DATE / ORA
# ─────────────────────────────────────────────────────────────────────────────

def parse_data(valore: str, riferimento: str, logger: logging.Logger):
    """
    Converte una stringa data in oggetto datetime.
    Prova diversi formati comuni italiani e ISO.
    """
    formati = [
        "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d",
        "%d/%m/%y", "%d-%m-%y",
        "%Y/%m/%d",
    ]
    valore = str(valore).strip()
    for fmt in formati:
        try:
            return datetime.strptime(valore, fmt)
        except ValueError:
            continue

    logger.warning(f"[{riferimento}] Formato data non riconosciuto: '{valore}' — riga ignorata.")
    return None


def parse_ora(valore: str) -> dt_time:
    """
    Converte una stringa ora in oggetto time.
    Se non riconosciuta, restituisce 09:00 come default.
    """
    if not valore or str(valore).strip() in ("", "nan", "None"):
        return dt_time(9, 0)

    formati = ["%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p"]
    valore = str(valore).strip()
    for fmt in formati:
        try:
            return datetime.strptime(valore, fmt).time()
        except ValueError:
            continue

    return dt_time(9, 0)


def combina_data_ora(data_obj: datetime, ora_obj: dt_time) -> datetime:
    """Combina un oggetto date/datetime con un oggetto time."""
    return datetime(
        data_obj.year, data_obj.month, data_obj.day,
        ora_obj.hour, ora_obj.minute, ora_obj.second
    )


# ─────────────────────────────────────────────────────────────────────────────
# INTERFACCIA OUTLOOK (win32com)
# ─────────────────────────────────────────────────────────────────────────────

def connetti_outlook(nome_calendario: str, logger: logging.Logger):
    """
    Connette a Outlook tramite COM e restituisce la cartella del calendario
    specificata nel config.
    """
    logger.info("Connessione a Microsoft Outlook in corso...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        namespace.Logon()
    except Exception as e:
        logger.error(f"Impossibile aprire Outlook: {e}")
        sys.exit(1)

    # Cerca il calendario per nome
    cartella_calendario = _cerca_calendario(namespace, nome_calendario, logger)

    if cartella_calendario is None:
        logger.error(
            f"Calendario '{nome_calendario}' non trovato in Outlook. "
            "Controlla il nome nel config.ini."
        )
        sys.exit(1)

    logger.info(f"Calendario trovato: '{cartella_calendario.Name}'")
    return cartella_calendario


def _cerca_calendario(namespace, nome: str, logger: logging.Logger):
    """
    Ricerca ricorsiva del calendario con il nome specificato.
    Controlla prima il calendario predefinito, poi le cartelle condivise.
    """
    # Calendario predefinito (olFolderCalendar = 9)
    try:
        calendario_default = namespace.GetDefaultFolder(9)
        if nome.lower() in ("default", "predefinito", calendario_default.Name.lower()):
            return calendario_default
    except Exception:
        pass

    # Scansione di tutti i negozi (account)
    def ricerca_in_cartella(cartella, nome_target):
        if cartella.Name.lower() == nome_target.lower():
            return cartella
        try:
            for sotto_cartella in cartella.Folders:
                risultato = ricerca_in_cartella(sotto_cartella, nome_target)
                if risultato:
                    return risultato
        except Exception:
            pass
        return None

    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
            risultato = ricerca_in_cartella(root, nome)
            if risultato:
                return risultato
        except Exception:
            continue

    return None


def costruisci_tag(riferimento: str) -> str:
    """Costruisce il tag univoco da inserire nel corpo dell'appuntamento."""
    return f"{TAG_PREFIX}{riferimento.strip()}{TAG_SUFFIX}"


def cerca_evento_esistente(calendario, riferimento: str, logger: logging.Logger):
    """
    Cerca nel calendario Outlook un evento che contenga il tag del riferimento
    nel campo Body. Usa la restrizione DASL per efficienza.
    Restituisce l'oggetto AppointmentItem o None.
    """
    tag = costruisci_tag(riferimento)

    try:
        items = calendario.Items
        items.IncludeRecurrences = False

        # Filtro testuale sul body tramite DASL
        filtro = f"@SQL=\"urn:schemas:httpmail:textdescription\" LIKE '%{tag}%'"
        items_filtrati = items.Restrict(filtro)

        for item in items_filtrati:
            if tag in (item.Body or ""):
                return item
    except Exception as e:
        logger.warning(
            f"[{riferimento}] Errore nella ricerca DASL, utilizzo ricerca lenta: {e}"
        )
        # Fallback: scansione lineare (più lenta ma affidabile)
        try:
            for item in calendario.Items:
                if tag in (item.Body or ""):
                    return item
        except Exception as e2:
            logger.error(f"[{riferimento}] Errore nella ricerca di fallback: {e2}")

    return None


def evento_e_aggiornato(item, titolo: str, inizio: datetime, descrizione: str, categoria: str) -> bool:
    """
    Confronta i dati dell'evento Outlook esistente con i dati Excel.
    Restituisce True se i dati sono cambiati (→ aggiornamento necessario).
    """
    # Confronto titolo
    if (item.Subject or "").strip() != titolo.strip():
        return True

    # Confronto data/ora inizio (tolleranza: 1 minuto)
    try:
        inizio_outlook = item.Start
        # pywintypes.datetime → datetime Python
        inizio_outlook_dt = datetime(
            inizio_outlook.year, inizio_outlook.month, inizio_outlook.day,
            inizio_outlook.hour, inizio_outlook.minute
        )
        if abs((inizio_outlook_dt - inizio).total_seconds()) > 60:
            return True
    except Exception:
        return True

    # Confronto categoria
    if categoria and (item.Categories or "").strip() != categoria.strip():
        return True

    return False


def crea_evento(calendario, riferimento: str, titolo: str, inizio: datetime,
                durata_minuti: int, descrizione: str, categoria: str,
                logger: logging.Logger) -> bool:
    """
    Crea un nuovo AppointmentItem nel calendario Outlook.
    Il tag [REF:...] viene aggiunto in fondo al corpo.
    """
    try:
        item = calendario.Items.Add(1)  # 1 = olAppointmentItem
        item.Subject = titolo
        item.Start = inizio
        item.Duration = durata_minuti

        # Corpo: descrizione + tag identificativo
        corpo = ""
        if descrizione and str(descrizione).strip() not in ("", "nan", "None"):
            corpo = str(descrizione).strip() + "\n\n"
        corpo += costruisci_tag(riferimento)
        item.Body = corpo

        if categoria and str(categoria).strip() not in ("", "nan", "None"):
            item.Categories = str(categoria).strip()

        item.ReminderSet = False  # Nessun promemoria di default
        item.Save()
        return True
    except Exception as e:
        logger.error(f"[{riferimento}] Errore nella creazione dell'evento: {e}")
        return False


def aggiorna_evento(item, riferimento: str, titolo: str, inizio: datetime,
                    durata_minuti: int, descrizione: str, categoria: str,
                    logger: logging.Logger) -> bool:
    """
    Aggiorna un AppointmentItem esistente con i nuovi dati dall'Excel.
    Preserva il tag [REF:...] nel corpo.
    """
    try:
        item.Subject = titolo
        item.Start = inizio
        item.Duration = durata_minuti

        # Ricostruisce il corpo preservando il tag
        corpo = ""
        if descrizione and str(descrizione).strip() not in ("", "nan", "None"):
            corpo = str(descrizione).strip() + "\n\n"
        corpo += costruisci_tag(riferimento)
        item.Body = corpo

        if categoria and str(categoria).strip() not in ("", "nan", "None"):
            item.Categories = str(categoria).strip()

        item.Save()
        return True
    except Exception as e:
        logger.error(f"[{riferimento}] Errore nell'aggiornamento dell'evento: {e}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# LOGICA PRINCIPALE DI SINCRONIZZAZIONE
# ─────────────────────────────────────────────────────────────────────────────

def sincronizza(df: pd.DataFrame, calendario, durata_default: int,
                logger: logging.Logger) -> dict:
    """
    Esegue la sincronizzazione riga per riga del DataFrame verso Outlook.
    Restituisce un dizionario con i contatori delle operazioni.
    """
    contatori = {"creati": 0, "aggiornati": 0, "invariati": 0, "errori": 0}

    totale = len(df)
    logger.info(f"Inizio sincronizzazione di {totale} righe...")

    for idx, riga in df.iterrows():
        riferimento = str(riga["Riferimento"]).strip()
        titolo = str(riga["Titolo"]).strip()
        data_str = str(riga["Data"]).strip()
        ora_str = str(riga.get("Ora", "")).strip()
        descrizione = str(riga.get("Descrizione", "")).strip()
        categoria = str(riga.get("Categoria", "")).strip()

        # Parse data
        data_obj = parse_data(data_str, riferimento, logger)
        if data_obj is None:
            contatori["errori"] += 1
            continue

        # Parse ora e combina
        ora_obj = parse_ora(ora_str)
        inizio = combina_data_ora(data_obj, ora_obj)

        # Cerca evento esistente
        evento = cerca_evento_esistente(calendario, riferimento, logger)

        if evento is None:
            # Evento non esiste → crea
            successo = crea_evento(
                calendario, riferimento, titolo, inizio,
                durata_default, descrizione, categoria, logger
            )
            if successo:
                logger.info(f"[{riferimento}] CREATO: '{titolo}' il {inizio:%d/%m/%Y %H:%M}")
                contatori["creati"] += 1
            else:
                contatori["errori"] += 1

        elif evento_e_aggiornato(evento, titolo, inizio, descrizione, categoria):
            # Evento esiste ma dati cambiati → aggiorna
            successo = aggiorna_evento(
                evento, riferimento, titolo, inizio,
                durata_default, descrizione, categoria, logger
            )
            if successo:
                logger.info(f"[{riferimento}] AGGIORNATO: '{titolo}' il {inizio:%d/%m/%Y %H:%M}")
                contatori["aggiornati"] += 1
            else:
                contatori["errori"] += 1

        else:
            # Evento esiste e dati identici → nessuna azione
            logger.debug(f"[{riferimento}] INVARIATO: '{titolo}'")
            contatori["invariati"] += 1

    return contatori


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # 1. Carica configurazione
    config = carica_configurazione("config.ini")

    percorso_excel = config.get("Excel", "percorso_file")
    nome_foglio = config.get("Excel", "nome_foglio", fallback="Scadenze")
    nome_calendario = config.get("Outlook", "nome_calendario", fallback="Calendario")
    durata_default = config.getint("Outlook", "durata_default_minuti", fallback=60)
    log_dir = config.get("Log", "cartella_log", fallback="logs")
    log_level = config.get("Log", "livello", fallback="INFO")

    # 2. Configura logging
    logger = configura_logging(log_dir, log_level)
    logger.info("=" * 60)
    logger.info("AVVIO SINCRONIZZAZIONE Excel → Outlook")
    logger.info("=" * 60)

    # 3. Leggi Excel
    df = leggi_excel(percorso_excel, nome_foglio, logger)

    # 4. Connetti a Outlook
    calendario = connetti_outlook(nome_calendario, logger)

    # 5. Sincronizza
    contatori = sincronizza(df, calendario, durata_default, logger)

    # 6. Riepilogo finale
    logger.info("-" * 60)
    logger.info("RIEPILOGO SINCRONIZZAZIONE:")
    logger.info(f"  Creati    : {contatori['creati']}")
    logger.info(f"  Aggiornati: {contatori['aggiornati']}")
    logger.info(f"  Invariati : {contatori['invariati']}")
    logger.info(f"  Errori    : {contatori['errori']}")
    logger.info("=" * 60)
    logger.info("Sincronizzazione completata.")


if __name__ == "__main__":
    main()
