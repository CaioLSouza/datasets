let
    Config = Excel.CurrentWorkbook(){[Name="tb_config"]}[Content],
    Caminho = Config{[campo="caminho_base"]}[valor],
    Fonte = Csv.Document(File.Contents(Caminho), [Delimiter=";", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    BaseCabecalho = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),
    Base = BaseCabecalho,
    Tipos = Table.TransformColumnTypes(Base, {
        {"onda", Int64.Type}, {"data_onda", type date}, {"nivel_dado", type text},
        {"resp_id", type text}, {"q_id", type text}, {"familia", type text}, {"safra", type text},
        {"opcao_id", type text}, {"opcao_pt", type text}, {"opcao_en", type text},
        {"valor_num", type number}, {"metrica", type text}, {"valor", type number},
        {"n", Int64.Type}, {"base", Int64.Type}, {"fonte", type text}
    }),
    Respostas = Table.SelectRows(Tipos, each [nivel_dado] = "respondente"),
    Publicado = Table.SelectRows(Tipos, each [nivel_dado] <> "respondente"),
    ChavesRaw = Table.Distinct(Table.SelectColumns(Respostas, {"onda", "q_id"})),
    PublicadoSemRaw = Table.NestedJoin(Publicado, {"onda", "q_id"}, ChavesRaw, {"onda", "q_id"}, "tem_raw", JoinKind.LeftAnti),

    RawOpcoes0 = Table.Group(Respostas,
        {"onda", "data_onda", "q_id", "familia", "safra", "opcao_id", "opcao_pt", "opcao_en"},
        {{"n", each Table.RowCount(_), Int64.Type}, {"base", each List.Count(List.Distinct([resp_id])), Int64.Type}}),
    RawOpcoes1 = Table.AddColumn(RawOpcoes0, "metrica", each "pct", type text),
    RawOpcoes2 = Table.AddColumn(RawOpcoes1, "valor", each if [base] = 0 then null else [n] / [base], type number),
    RawOpcoes3 = Table.AddColumn(RawOpcoes2, "fonte", each "calculado", type text),

    RawMedias0 = Table.Group(Respostas,
        {"onda", "data_onda", "q_id", "familia", "safra"},
        {{"n", each List.Count(List.RemoveNulls([valor_num])), Int64.Type},
         {"valor", each let x = List.RemoveNulls([valor_num]) in if List.Count(x) = 0 then null else List.Average(x), type number}}),
    RawMedias1 = Table.SelectRows(RawMedias0, each [valor] <> null),
    RawMedias2 = Table.AddColumn(RawMedias1, "opcao_id", each "", type text),
    RawMedias3 = Table.AddColumn(RawMedias2, "opcao_pt", each "Média", type text),
    RawMedias4 = Table.AddColumn(RawMedias3, "opcao_en", each "Average", type text),
    RawMedias5 = Table.AddColumn(RawMedias4, "base", each [n], Int64.Type),
    RawMedias6 = Table.AddColumn(RawMedias5, "metrica", each "media", type text),
    RawMedias7 = Table.AddColumn(RawMedias6, "fonte", each "calculado", type text),

    PublicadoOpcoes = Table.SelectRows(PublicadoSemRaw, each [metrica] = "pct"),
    PublicadoMedias = Table.SelectRows(PublicadoSemRaw, each [metrica] = "media"),
    Colunas = {"onda", "data_onda", "q_id", "familia", "safra", "opcao_id", "opcao_pt", "opcao_en", "metrica", "valor", "n", "base", "fonte"},
    Saida = Table.Combine({
        Table.SelectColumns(RawOpcoes3, Colunas),
        Table.SelectColumns(RawMedias7, Colunas),
        Table.SelectColumns(PublicadoOpcoes, Colunas),
        Table.SelectColumns(PublicadoMedias, Colunas)
    }),
    Ordenada = Table.Sort(Saida, {{"onda", Order.Ascending}, {"q_id", Order.Ascending}, {"metrica", Order.Ascending}, {"opcao_id", Order.Ascending}})
in
    Ordenada
