# Build e distribuzione EXE (Windows)

Questa repository usa una GitHub Action su `windows-latest` per generare l'eseguibile `CampionaturaSAP.exe` con PyInstaller.

## 1) Build automatico su GitHub Actions

Workflow: `.github/workflows/build-windows-exe.yml`

La pipeline:
1. Installa Python 3.11
2. Installa le dipendenze da `requirements.txt`
3. Installa `pyinstaller`
4. Esegue:
   ```bash
   pyinstaller --onefile --name CampionaturaSAP sap_sampling.py
   ```
5. Pubblica l'artifact `CampionaturaSAP-windows` (file `CampionaturaSAP.exe`)

## 2) Come scaricare l'artifact

1. Apri il repository su GitHub.
2. Vai in **Actions**.
3. Seleziona il run del workflow **Build Windows EXE**.
4. Nella sezione **Artifacts**, scarica `CampionaturaSAP-windows`.
5. Estrai il file zip: otterrai `CampionaturaSAP.exe`.

## 3) Uso da USB (INPUT/OUTPUT)

1. Copia su chiavetta USB:
   - `CampionaturaSAP.exe` (scaricato dagli artifact), oppure `RUN.bat` + sorgenti Python.
   - la cartella `INPUT/`
   - la cartella `OUTPUT/`
2. Inserisci i file Excel di input in `INPUT/`.
3. Avvia `CampionaturaSAP.exe` (o `RUN.bat`).
4. Verifica i risultati in `OUTPUT/`.

> `RUN.bat` usa prima `CampionaturaSAP.exe` (root o `dist/`), e in fallback avvia `python sap_sampling.py`.
