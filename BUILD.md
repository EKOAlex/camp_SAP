# Build dell'eseguibile (Windows)

Queste istruzioni usano **PyInstaller** per creare un `.exe` a partire da `sap_sampling.py`.

## Prerequisiti
1. Python 3.11+ installato e disponibile in PATH.
2. Dipendenze installate:
   ```bash
   python -m pip install -r requirements.txt
   python -m pip install pyinstaller
   ```

## Build
Esegui dalla root del repo:
```bash
pyinstaller --onefile --name sap_sampling sap_sampling.py
```

L'eseguibile finale si troverà in `dist/sap_sampling.exe`.

## Uso
1. Crea/usa le cartelle `INPUT` e `OUTPUT`.
2. Inserisci i file Excel in `INPUT` (con "ZSD032" e "ZPP04" nel nome).
3. Avvia `sap_sampling.exe` oppure `RUN.bat`.
