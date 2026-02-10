let
    // ========================================
    // CONFIGURAZIONE PERCORSI (relativi al file Excel template)
    // ========================================
    RootFolder = Text.From(Excel.CurrentWorkbook(){[Name="cfg_root"]}[Content]{0}[Column1]),
    InputFolder = RootFolder & "\\INPUT",

    // ========================================
    // FUNZIONE DI SUPPORTO: apre il primo file che contiene un token nel nome
    // ========================================
    GetFirstFileByToken = (folder as text, token as text) as binary =>
        let
            Files = Folder.Files(folder),
            Filtered = Table.SelectRows(Files, each Text.Contains(Text.Upper([Name]), Text.Upper(token)) and [Extension] = ".xlsx"),
            FirstFile = if Table.RowCount(Filtered) = 0
                then error "Nessun file trovato in " & folder & " con token " & token
                else Filtered{0}[Content]
        in
            FirstFile,

    // ========================================
    // LETTURA DATI INPUT
    // ========================================
    ZSD032Bin = GetFirstFileByToken(InputFolder, "ZSD032"),
    ZPP04Bin = GetFirstFileByToken(InputFolder, "ZPP04"),

    ZSD032WB = Excel.Workbook(ZSD032Bin, null, true),
    ZPP04WB = Excel.Workbook(ZPP04Bin, null, true),

    ZSD032Data = ZSD032WB{0}[Data],
    ZPP04Data = ZPP04WB{0}[Data],

    ZSD032Headers = Table.PromoteHeaders(ZSD032Data, [PromoteAllScalars=true]),
    ZPP04Headers = Table.PromoteHeaders(ZPP04Data, [PromoteAllScalars=true]),

    // ========================================
    // NORMALIZZAZIONE TIPI
    // ========================================
    Ordini = Table.TransformColumnTypes(
        ZSD032Headers,
        {
            {"Destinatario merci", type text},
            {"Materiale", type text},
            {"Descrizione", type text},
            {"Requested Date", type date},
            {"Qtà da Spedire", type number}
        },
        "it-IT"
    ),

    Materiali = Table.TransformColumnTypes(
        ZPP04Headers,
        {
            {"Materiale", type text},
            {"Plant-Specific Material Status", type text},
            {"T.m.", type text}
        },
        "it-IT"
    ),

    // ========================================
    // LOGICA BUSINESS
    // - solo materiali con stato "A"
    // - per materiale scegli l'ordine con Requested Date più vicina (data minima)
    // - is_kit = TRUE se T.m. contiene "t.m" (match case-insensitive)
    // ========================================
    MaterialiA = Table.SelectRows(Materiali, each [#"Plant-Specific Material Status"] = "A"),

    JoinOrdiniMateriali = Table.NestedJoin(
        Ordini,
        {"Materiale"},
        MaterialiA,
        {"Materiale"},
        "Mat",
        JoinKind.Inner
    ),

    Expanded = Table.ExpandTableColumn(JoinOrdiniMateriali, "Mat", {"T.m."}, {"T.m."}),

    CleanDates = Table.SelectRows(Expanded, each [#"Requested Date"] <> null),

    SortedByDate = Table.Sort(CleanDates, {{"Materiale", Order.Ascending}, {"Requested Date", Order.Ascending}}),

    EarliestPerMaterial = Table.Distinct(SortedByDate, {"Materiale"}),

    WithIsKit = Table.AddColumn(
        EarliestPerMaterial,
        "is_kit",
        each Text.Contains(Text.Lower(Text.From([#"T.m."])), "t.m"),
        type logical
    ),

    Piano = Table.SelectColumns(
        WithIsKit,
        {
            "Destinatario merci",
            "Materiale",
            "Descrizione",
            "Requested Date",
            "Qtà da Spedire",
            "is_kit"
        }
    ),

    PianoSorted = Table.Sort(Piano, {{"Destinatario merci", Order.Ascending}, {"Requested Date", Order.Ascending}})
in
    PianoSorted
