"""
Modulo per generare un piano di campionatura dei nuovi prodotti
partendo dalle estrazioni SAP (ZSD032 e ZPP04).

Il programma legge i due file Excel forniti da SAP, identifica i
codici materiale con stato "A6" nel file ZPP04 (che rappresenta
nuovi prodotti) e seleziona, per ciascun codice materiale, l’ordine
di vendita con la data di consegna più prossima presente nel file
ZSD032.  Nel piano risultante vengono indicati il cliente
(destinatario merci), il codice materiale, la descrizione, la data
di consegna (richiesta) e la quantità da spedire.  È inoltre
segnalato se il materiale è un kit sulla base del campo "T.m." del
file ZPP04 (valori che contengono "KIT").

L’output viene salvato in un file Excel denominato
``piano_campionatura.xlsx`` nella stessa cartella di lavoro.

Usage:
    python sap_sampling.py ZSD032.xlsx ZPP04.xlsx

Sono necessari i pacchetti ``pandas`` e ``openpyxl``.  Per
l’esecuzione da chiavetta USB su sistemi Windows è possibile
generare un eseguibile con PyInstaller (vedere la documentazione
nella sezione ``if __name__ == '__main__'``).
"""

import sys
from pathlib import Path
import pandas as pd


def genera_piano_campionatura(path_zsd032: Path, path_zpp04: Path) -> pd.DataFrame:
    """Crea il piano di campionatura combinando i due file SAP.

    Parameters
    ----------
    path_zsd032 : Path
        Percorso al file di estrazione ZSD032 contenente gli ordini di vendita.
    path_zpp04 : Path
        Percorso al file di estrazione ZPP04 contenente l'anagrafica dei
        materiali e i relativi stati.

    Returns
    -------
    pd.DataFrame
        DataFrame con le colonne: ``Destinatario merci``, ``Materiale``,
        ``Descrizione``, ``Requested Date``, ``Qtà da Spedire``, ``is_kit``.
    """
    # Carica l'estrazione ordini (ZSD032). La prima (e unica) sheet è "Data".
    ordini = pd.read_excel(path_zsd032, sheet_name=0)

    # Carica l'estrazione materiali (ZPP04).
    materiali = pd.read_excel(path_zpp04, sheet_name=0)

    # Seleziona solo i materiali con stato A6 (nuovi prodotti).
    if 'Plant-Specific Material Status' not in materiali.columns:
        raise KeyError(
            "Colonna 'Plant-Specific Material Status' mancante nel file ZPP04."
        )
    materiali_a6 = materiali[materiali['Plant-Specific Material Status'] == 'A6'].copy()
    # Crea un campo stringa per unire sulla colonna Materiale.  In alcuni
    # sistemi il campo potrebbe essere trattato come numerico o stringa; per
    # evitare problemi convertiamo entrambi a stringa senza zeri iniziali.
    ordini['Materiale_str'] = ordini['Materiale'].astype(str)
    materiali_a6['Materiale_str'] = materiali_a6['Materiale'].astype(str)

    # Esegui la fusione (join) tra ordini e materiali A6 sulla chiave
    # ``Materiale_str``. Utilizziamo un inner join per mantenere solo i
    # materiali che sono effettivamente presenti negli ordini.
    merged = ordini.merge(
        materiali_a6[['Materiale_str', 'T.m.']], on='Materiale_str', how='inner'
    )

    # Converti la colonna di data di consegna in datetime.  Si assume che
    # ``Requested Date`` rappresenti la data di consegna richiesta (scadenza).
    merged['Requested Date'] = pd.to_datetime(merged['Requested Date'], errors='coerce')
    # Rimuovi eventuali righe con data mancante.
    merged = merged.dropna(subset=['Requested Date'])

    # Seleziona l'ordine con data più prossima (la minima) per ciascun
    # codice materiale.  ``idxmin`` restituisce l'indice della prima
    # occorrenza del valore minimo in ciascun gruppo.
    idx_min_dates = merged.groupby('Materiale_str')['Requested Date'].idxmin()
    selezione = merged.loc[idx_min_dates].copy()

    # Prepara il campo ``is_kit`` per indicare se è un kit.  Sono
    # considerati kit i materiali il cui campo ``T.m.`` contiene
    # ``KIT`` (indipendentemente dal maiuscolo/minuscolo).  In assenza
    # della colonna ``T.m.``, il valore sarà False.
    def is_kit_val(val: str) -> bool:
        return isinstance(val, str) and ('KIT' in val.upper())

    selezione['is_kit'] = selezione['T.m.'].apply(is_kit_val)

    # Seleziona e rinomina le colonne rilevanti del piano.
    piano = selezione[[
        'Destinatario merci',
        'Materiale',
        'Descrizione',
        'Requested Date',
        'Qtà da Spedire',
        'is_kit',
    ]].copy()

    # Ordina il piano per cliente (destinatario merci) e data di consegna.
    piano = piano.sort_values(['Destinatario merci', 'Requested Date'])
    return piano.reset_index(drop=True)


def main(argv: list[str] | None = None) -> None:
    """Punto di ingresso del programma da riga di comando.

    Accetta due argomenti: percorso al file ZSD032 e percorso al file
    ZPP04.  Genera il piano di campionatura e lo salva in un file
    ``piano_campionatura.xlsx``.
    """
    argv = argv or sys.argv[1:]
    if len(argv) != 2:
        print(
            'Uso: python sap_sampling.py <path_ZSD032.xlsx> <path_ZPP04.xlsx>',
            file=sys.stderr,
        )
        sys.exit(1)
    path_zsd, path_zpp = map(Path, argv)
    if not path_zsd.exists():
        raise FileNotFoundError(f'File ZSD032 non trovato: {path_zsd}')
    if not path_zpp.exists():
        raise FileNotFoundError(f'File ZPP04 non trovato: {path_zpp}')
    # Genera il piano.
    piano_df = genera_piano_campionatura(path_zsd, path_zpp)
    # Scrive su disco l'output come file Excel.  Utilizziamo
    # l'estensione xlsx e la libreria openpyxl.
    output_path = path_zsd.parent / 'piano_campionatura.xlsx'
    piano_df.to_excel(output_path, index=False)
    print(f'Piano campionatura salvato in {output_path}')


if __name__ == '__main__':
    main()