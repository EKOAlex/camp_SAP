# Template Excel-only (Power Query)

Questo template sostituisce completamente l'approccio `.exe` / Python.

## Obiettivo
- **INPUT**: cartella `INPUT\` con 2 file Excel (uno ZSD032 e uno ZPP04).
- **OUTPUT**:
  - tabella Excel chiamata **`piano`**;
  - fogli separati per cliente, generati dalla tabella `piano`.

## Prerequisiti nel file Excel template
1. Crea un nuovo file, ad esempio `Template_Campionatura.xlsx` nella root del progetto.
2. In un foglio di servizio (es. `CONFIG`), inserisci nella cella `A1` il percorso assoluto della cartella progetto (es. `C:\camp_SAP`).
3. Trasforma `A1` in tabella (Ctrl+T) e assegna alla tabella il nome **`cfg_root`**.

## Query Power Query
1. Apri Excel > **Dati** > **Recupera dati** > **Da altre origini** > **Query vuota**.
2. Apri **Editor avanzato**.
3. Copia/incolla il contenuto di `TEMPLATE/piano_query.m`.
4. Salva la query con nome **`piano`**.
5. **Chiudi e carica in...** > **Tabella** in un nuovo foglio.

## Logica implementata nella query
- Usa solo record con `Plant-Specific Material Status = "A"`.
- Per ogni `Materiale` seleziona l'ordine con `Requested Date` più vicina (data minima).
- Colonna `is_kit = TRUE` se `T.m.` contiene la stringa `t.m` (case-insensitive).

## Generazione fogli cliente (senza macro)
Per ottenere un foglio per ciascun cliente:
1. Seleziona la tabella `piano`.
2. **Inserisci** > **Tabella Pivot** (nuovo foglio).
3. In campi Pivot:
   - Filtro: `Destinatario merci`
   - Righe: `Materiale`, `Descrizione`, `Requested Date`
   - Valori: `Qtà da Spedire` (somma)
4. Dalla Pivot: **Analizza tabella pivot** > **Opzioni** > **Mostra pagine filtro report...**
5. Scegli `Destinatario merci`: Excel crea automaticamente un foglio per ogni cliente.

> Dopo ogni aggiornamento della query, ripeti “Mostra pagine filtro report” per rigenerare i fogli cliente aggiornati.
