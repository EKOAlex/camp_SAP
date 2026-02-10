# Campionatura SAP - soluzione Excel-only (Power Query)

Questa versione usa **solo Excel**: nessun `.exe`, nessun requisito Python.

## Struttura cartelle
- `INPUT\`: inserire i 2 file Excel SAP (`ZSD032...xlsx` e `ZPP04...xlsx`)
- `OUTPUT\`: opzionale per salvataggi manuali/export
- `TEMPLATE\`: contiene il template della query Power Query
  - `piano_query.m`
  - `POWER_QUERY_TEMPLATE.md`

## Procedura iniziale (una sola volta)
1. Apri un nuovo file Excel (es. `Template_Campionatura.xlsx`) nella root del progetto.
2. Crea un foglio `CONFIG`.
3. In `CONFIG!A1` inserisci il path assoluto della root progetto (es. `C:\camp_SAP`).
4. Trasforma `A1` in Tabella (Ctrl+T) e rinominala **`cfg_root`**.
5. Vai su **Dati > Recupera dati > Da altre origini > Query vuota**.
6. Apri **Editor avanzato** e incolla il contenuto di `TEMPLATE/piano_query.m`.
7. Rinomina la query in **`piano`**.
8. **Chiudi e carica in... > Tabella** in un nuovo foglio.

## Logica del piano
La query applica questa logica:
- tiene solo materiali con stato `A`;
- per ogni materiale seleziona l'ordine con `Requested Date` più vicina (la minima);
- imposta `is_kit = TRUE` se `T.m.` contiene `t.m` (senza distinzione maiuscole/minuscole).

## Come aggiornare i dati (operatività)
1. Sostituisci i file in `INPUT\` con i nuovi export SAP.
2. Apri `Template_Campionatura.xlsx`.
3. Clicca **Dati > Aggiorna tutto**.
4. Attendi il refresh della query `piano`.
5. (Se usi i fogli per cliente con Pivot) rigenera i fogli con:
   - **Analizza tabella pivot > Opzioni > Mostra pagine filtro report...**
   - seleziona `Destinatario merci`.

## Output
- Tabella principale: **`piano`**.
- Fogli per cliente: creati dalla Pivot con “Mostra pagine filtro report”.
