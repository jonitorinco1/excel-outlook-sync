@echo off
REM ============================================================
REM run_sync.bat — Avvio sincronizzazione Excel → Outlook
REM ============================================================
REM Questo file viene usato per eseguire lo script tramite
REM Task Scheduler di Windows o manualmente da riga di comando.
REM
REM CONFIGURAZIONE (modifica i percorsi qui sotto):
REM   PYTHON_EXE  = percorso all'eseguibile Python nel tuo ambiente
REM   SCRIPT_DIR  = cartella radice del progetto
REM ============================================================

REM --- Percorso all'eseguibile Python ---
REM Se usi un virtualenv, punta all'exe dentro .venv\Scripts\python.exe
REM Es.: SET PYTHON_EXE=C:\Users\TuoNome\project\excel-outlook-sync\.venv\Scripts\python.exe
SET PYTHON_EXE=python

REM --- Cartella del progetto (dove si trova sync.py) ---
SET SCRIPT_DIR=%~dp0

REM --- Spostati nella cartella del progetto ---
cd /d "%SCRIPT_DIR%"

REM --- Esegui lo script di sincronizzazione ---
"%PYTHON_EXE%" sync.py

REM --- Pausa solo se eseguito manualmente (non da Task Scheduler) ---
IF "%1"=="manual" PAUSE

exit /b %ERRORLEVEL%
