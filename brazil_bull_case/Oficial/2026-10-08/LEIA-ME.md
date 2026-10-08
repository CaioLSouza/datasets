# Modelos oficiais até 08/out/2026

Valem até nova versão. Os números ainda precisam passar pelo Fernando e pelo compliance antes de ir para clientes.

## Arquivos

| Arquivo | O que é | Origem |
|---|---|---|
| `Ibov_multiplos_bottomup_oficial_2026-10-08.xlsx` | Ibovespa por caixinha pelos 3 métodos: P/L e EV/EBITDA de 10 anos (±1 desvio-padrão) e bottom-up (Ke ±200 bp). Cobre todas as caixinhas. | `Ibov_valuation_by_box_model_XP.xlsx` da pasta do projeto, gerado por `scripts/build_model.py` |
| `Financials_DCF_oficial_2026-10-08.xlsx` | DCF de Financials (FCFE / DDM de banco), com a lógica dos analistas sobre o consenso. | `Financials_DDM_XP_v6.xlsx`, gerado por `scripts/build_fin_ddm.py` |
| `scripts/` | Scripts que geram os dois arquivos, na versão desta data. | Cópia de `scripts/` |

## Premissas do DCF de Financials

- **Conteúdo:** consenso Bloomberg BEst de 2026–28 dos membros (lucro e ROE), somado na caixinha. A projeção é feita só no nível da caixinha.
- **Lógica (analistas):**
  - ROE fixo no de 2028;
  - crescimento caindo até g;
  - fluxo ao acionista = lucro − aumento do patrimônio;
  - Gordon sobre o último fluxo, descontado pelo fator do último ano.
- **Ke nominal:** (NTN-B + IPCA 4%) × 0,85 + beta × ERP. O ERP é de 5,5% e igual para todas as caixinhas.
- **Escolhas casa × analistas:**
  - beta 1,15 da caixinha (casa);
  - g 6% = IPCA + PIB real (analistas);
  - 4 anos explícitos (casa);
  - perpetuidade sem período extra (analistas).
- **Cenários:** juro real 8,5% / hoje / 5,5%; choque de lucro ±20% × beta de lucro normalizado (Financials ±17,4%).
- **Resultado:** −36% / −10% / +20% (bear / base / bull), valor no fim de 2027 contra os pontos de 05/out.

## Atenção

- **Cópia do modelo de múltiplos:** é a do projeto. A cópia em `Downloads` estava aberta no Excel e não pôde ser lida. Se você editou aquela, salve, feche e substitua.
- **Financials no Excel:** a cópia é a versão salva em disco. Se houver alterações não salvas no Excel, elas não entraram.
