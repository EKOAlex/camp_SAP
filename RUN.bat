@echo off
setlocal
cd /d "%~dp0"

if exist "CampionaturaSAP.exe" (
  echo Avvio eseguibile locale CampionaturaSAP.exe
  "CampionaturaSAP.exe"
) else if exist "dist\CampionaturaSAP.exe" (
  echo Avvio eseguibile buildato in dist\CampionaturaSAP.exe
  "dist\CampionaturaSAP.exe"
) else (
  echo Nessun exe trovato (artifact GitHub Actions non presente in locale).
  echo Avvio via Python: sap_sampling.py
  python sap_sampling.py
)

pause
