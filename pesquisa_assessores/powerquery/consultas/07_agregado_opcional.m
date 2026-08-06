// Uma linha por onda x pergunta x alternativa, de jul/2023 até hoje.
// Para tabela dinâmica, recorte novo, conferência.
//
// Use SEMPRE a coluna `pct` — é ela que respeita o congelamento:
//   fonte = "publicado"  -> o número que já foi ao ar, intocado
//   fonte = "calculado"  -> onda nova, calculada do bruto
// A coluna `pct_calculado` é só auditoria. Não aponte gráfico para ela.
let
    Arquivo = PastaBases & "\PA Base Historica.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="agregado", Kind="Sheet"]}[Data],
    Cab     = Table.PromoteHeaders(Aba, [PromoteAllScalars=true]),
    Tipado  = Table.TransformColumnTypes(Cab, {
        {"onda", Int64.Type}, {"q_id", type text}, {"familia", type text},
        {"safra", type text}, {"pergunta_pt", type text},
        {"pergunta_en", type text}, {"opcao_id", type text},
        {"opcao_pt", type text}, {"opcao_en", type text},
        {"chave", type text}, {"ordem_pergunta", Int64.Type},
        {"ordem_opcao", Int64.Type}, {"n", Int64.Type},
        {"base", Int64.Type}, {"pct", type number},
        {"pct_calculado", type number}, {"fonte", type text}})
in
    Tipado
