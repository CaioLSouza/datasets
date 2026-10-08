# Briefing: valuation do Ibovespa por box (bear / base / bull)

Base do report "Brazil Bull Case". Este arquivo resume o estudo feito em 05–06/out/2026 para um cliente.

- **Fonte original:** `Equity Strategy Dashboard\_pedidos\2026-10-05_ibov_valuation_groups\`
  - Modelo completo, com fórmulas: `Ibov_valuation_by_box_2026-10-05.xlsx`
  - Scripts: `scripts\bx_model.py`, `ray_table.py`, `charts10.py`, `sens_org.py`
- **Cópias nesta pasta:** `inputs\`, com as tabelas finais (v6/v3), P/E e ERP 10 anos, sensibilidade a Ke e os PNGs.
- **Preços:** fechamento de 05/out/2026. O Ibovespa fechou em 208.432 pts, depois do rali pós-eleição. A meta da casa para o fim de 2026 é 200 mil.

## 1. Os quatro boxes (GICS ajustado XP)

| Box | Peso no Ibov | Pts | P/E 12m fwd hoje | Média 10a (±1σ) | ERP hoje | ERP média 10a |
|---|---|---|---|---|---|---|
| Financials | 29% | 60.927 | 9,8x | 8,6x (7,0–10,1) | 3,4% | 6,8% |
| Defensives (UT/CS/HC/Telecom) | 27% | 55.822 | 14,6x | 13,8x (11,4–16,2) | 0,1% | 2,2% |
| Cyclicals (CD/IN/IT/RE) | 13% | 26.271 | 12,9x | 17,9x (13,2–22,6) | 1,0% | 0,7% |
| Commodities (EN/MT) | 31% | 65.411 | 6,4x | 7,4x (4,7–10,0) | 8,8% | 9,9% |
| **Ibovespa** | 100% | 208.432 | **9,4x** | 9,4x (7,2–11,5) | 3,9% | 6,0% |
| *memo:* Cyclicals ex-WEG & Embraer | 8% | 16.067 | 10,8x | | | |
| *memo:* WEG & Embraer | 5% | 10.205 | 18,9x | | | |

O ERP é o earnings yield 12m fwd menos o juro real de 5 anos (NTN-B), mesma definição do Chart 26 do Ray.

## 2. Metodologia final (Caio, 06/out)

**Cenários macro, contra uma Selic de 13,75% hoje:**
- **Bear:** populista moderado, Selic 13,20%.
- **Base:** Selic 11,50%. A curva LTN de 05/out já precifica ~11,4–11,9% no fim de 2027, então o base ≈ a curva de hoje.
- **Bull:** reformista, Selic 9,50%.

**Três métodos por box:**
1. **P/E 12m fwd:** média de 10 anos (out/16–set/26), com −1 / 0 / +1σ.
2. **EV/EBITDA 12m fwd:** mesma regra; não se aplica a Financials.
3. **Bottom-up:** preços-alvo 12m da XP, da COMP SHEET, sem ÷1,15.
   - Bear e bull são reprecificados com Ke +200bp / −200bp, usando a sensibilidade do TP a Ke informada pelos analistas.
   - Há um cap de ±40% por ação, que trava WEGE3 e CSNA3.
   - 200bp ≈ 0,8σ das variações de 12 meses do pré 5a, ou ~90bp de Selic + ~110bp de prêmio fiscal.

**Ajuste de lucro por juros (dentro dos múltiplos, não é um método separado):**
- O EPS do Ibov sobe 2,07% a cada −100bp de Selic (regressão com 6m de defasagem e controles).
- O efeito é repartido entre os boxes pelo beta de retorno ao pré 2a, normalizado para o Ibov.
- Variação de EPS no Ibov: +1,1% / +4,7% / +8,8%.

**Média:** simples entre os métodos; Financials usa só P/E e bottom-up. O Ibov é a soma dos boxes.

## 3. Resultado: upside vs fechamento de 05/out

| | Bear | Base | Bull |
|---|---|---|---|
| **Ibovespa (pts)** | **174.414** | **229.232** | **294.832** |
| Ibovespa (média) | −16,3% | +10,0% | +41,5% |
| P/E (7,3x / 9,8x / 12,3x) | −21,3% | +9,4% | +42,9% |
| EV/EBITDA (soma dos boxes) | −17,8% | +13,2% | +47,0% |
| Bottom-up (Ke +200 / 0 / −200bp) | −10,3% | +7,4% | +33,8% |
| Financials | −26,3% | −7,5% | +17,6% |
| Defensives | −13,5% | +7,1% | +31,5% |
| Cyclicals | +1,6% | +39,7% | +84,7% |
| Commodities | −16,7% | +16,9% | +54,8% |
| *memo:* Cyclicals ex-WEG & EMB | +13,4% | +62,8% | +121,5% |
| *memo:* WEG & Embraer | −25,6% | +25,8% | +82,2% |

**Contribuição para o upside do Ibov, em pp (bear / base / bull):**

| Box | Bear | Base | Bull |
|---|---|---|---|
| Financials | −7,7 | −2,2 | +5,2 |
| Defensives | −3,6 | +1,9 | +8,4 |
| Cyclicals | +0,2 | +5,0 | +10,7 |
| Commodities | −5,2 | +5,3 | +17,2 |

## 4. Estudos de apoio (`studies\`)

**Performance vs juros** (semanal desde 2012, pré 2a, controles BCOM + USDBRL, Newey-West). Retorno para cada −100bp de pré 2a:

| Box | Retorno absoluto | Relativo ao Ibov |
|---|---|---|
| Ibov | +2,9% (t 7,0) | — |
| Financials | +3,6% | +0,7pp |
| Defensives | +2,8% | −0,1pp |
| Cyclicals | +4,3% | +1,4pp |
| Cyclicals ex-WEG/EMB | +5,2% | **+2,3pp** (t 6,7) |
| Commodities | +1,8% | −1,1pp |
| WEG & EMB | +1,3% | −1,6pp |

A rotação começa antes do primeiro corte de Selic (event study disponível desde 2006).

**EPS vs juros:**
- Só Defensives é robusto: −1,7% de EPS por +1pp de Selic (t −4).
- Em Commodities, o beta reflete o ciclo de commodities.
- Em Financials e Cyclicals, não é significativo.

**Sensibilidade a Ke (analistas), Ke −100bp → TP:**
- Fin +12,9%, Def +9,3%, Cyc +17,4% (ex-WEG/EMB +9,3%), Comm +10,2%, Ibov +11,5%.
- Cobertura de 96% do Ibov.

## 5. Mensagens que saem dos números

- **O box que faz o bull case é Cyclicals ex-WEG/EMB.** É o único com ERP acima da média de 10 anos, enquanto o P/E está 1σ abaixo da média. É também o box mais sensível a juros (+2,3pp vs Ibov por −100bp).
- **Financials já estão caros.** O P/E de 9,8x está acima de +1σ e o upside no base é negativo (−7,5%). Foi o box que mais andou no rali.
- **Defensives também estão esticados.** O ERP está perto de zero (0,1% vs 2,2% de média), então dependem de queda de juro real.
- **No bull, Commodities respondem por 17pp dos 41,5% de upside.** O P/E de 6,4x está abaixo da média de 7,4x.
- **O Ibov a 9,4x está exatamente na média de 10 anos.** O ERP de 3,9% está perto de −1σ: o rali já consumiu parte do prêmio.

## 6. Pendências antes de publicar

- [ ] Atualizar preços e Selic/curva de 05/out para a data de corte do report.
- [ ] Corrigir valores velhos da Bloomberg no `bx_panel20`, onde o BDH repete o último valor: NATU3, GOAU4, BHIA3, PCAR3.
  - O dashboard já corrigiu. O efeito nas médias é pequeno: Commodities 7,37x → 7,49x, Ibov 9,36x → 9,42x.
- [ ] Sensibilidades a Ke faltantes: BBSE3, PSSA3, CSAN3, HAPV3, BEEF3. Hoje usam a média do subgrupo.
- [ ] Checar os outliers WEGE3 (+38%) e RAPT4 (+40%) na sensibilidade a Ke.
- [ ] Os números novos, que não estão no Ray, passam por Fernando/compliance antes de ir ao cliente.
- [ ] Os números do house model (`DCF Ibov 2026_Out.xlsx`), com 9,5x P/E × 2026E e meta de 200k, não batem 1:1 com a abordagem 12m fwd. É preciso explicar a ponte no report.
