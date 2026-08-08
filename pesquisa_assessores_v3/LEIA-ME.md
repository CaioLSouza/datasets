# Pesquisa de Assessores XP — atualização mensal

## O mês inteiro

```
1. Exportar o Forms      ──►  input_forms\
2. Duplo clique em       ──►  rodar.bat
3. Abrir PA Charts.xlsx  ──►  Dados > Atualizar Tudo
```

Os gráficos das perguntas recorrentes se atualizam sozinhos. **A exceção é a
pergunta do mês**: ela muda de forma a cada edição, então é o único gráfico que
você remonta — os números já vêm apurados na aba `q_mes`.

Duas coisas seguem manuais, e por motivo:

| Manual | Por quê |
|---|---|
| Tradução da pergunta do mês para inglês | Ninguém traduz melhor que você. Coluna `rotulo_en` da aba `q_mes`. |
| Fechamento do Ibovespa da capa | Não sai da pesquisa. Uma linha por mês em `ibovespa.csv`; o log avisa qual falta. |

---

## Como as peças se encaixam

```
input_forms\*.xlsx          o export do Forms, como ele sai
        │
        ▼   atualizar.py
   ┌────┴──────────────┬──────────────────────┐
   ▼                   ▼                      ▼
PA Base.xlsx    PA Charts Data.xlsx     bases\charts\*.csv
histórico       as mesmas tabelas       a fonte do Power Query
completo,       numa pasta só, para     (um arquivo por tabela)
formato longo   você olhar                     │
                                               ▼  Power Query
                                       PA Charts.xlsx
                                       suas tabelas + seus gráficos
```

O Python não sabe onde nenhum gráfico mora. Ele entrega **tabelas**; o Excel
cuida do resto. É isso que faz "Atualizar Tudo" bastar.

**Por que a fonte é CSV e não as abas do xlsx:** cada consulta lê só o seu
arquivo. Lendo do xlsx, as 27 consultas reparseavam a pasta de trabalho inteira
uma a uma e o Atualizar Tudo levava **143 segundos**. Com um CSV por tabela caiu
para **35**. A `PA Charts Data.xlsx` continua sendo gerada, mas só para você
folhear — ninguém a consulta.

Os CSVs saem em formato fixo: UTF-8 com BOM, vírgula, ponto decimal, data ISO. As
consultas declaram `Culture="en-US"` para casar com isso, então a configuração
regional da máquina não muda o resultado.

---

## As tabelas

Cada tabela vira uma consulta Power Query na `PA Charts.xlsx`, carregada como
Tabela na planilha do mesmo nome.

| Aba | Uma linha por | Colunas |
|---|---|---|
| `d_<pergunta>` | alternativa | `ordem`, `alternativa_id`, `rotulo_pt`, `rotulo_en`, `atual`, `anterior`, `delta` |
| `s_<pergunta>` | onda | `onda`, `data`, e uma coluna por alternativa |
| `medias` | onda | `sentimento_media`, `ibovespa_media`, `respondentes` |
| `capa` | onda | as três séries de intenção + `ibovespa` |
| `meta` | onda | `data`, `respondentes`, `regime` |
| `corrente` | — | uma linha só: a onda do report, `mes_pt`, `mes_en` |
| `q_mes` | alternativa | a pergunta do mês |

As 11 perguntas recorrentes: `regiao`, `alocacao_rv`, `proximos_meses`,
`classes_ativos`, `pct_internacional`, `interesse_internacional`,
`riscos_bolsa`, `setores`, `sentimento`, `ibovespa_alvo`, `apetite_risco`.

### As duas coisas que economizam trabalho

**`d_` e `s_` servem os dois decks.** Há duas colunas de rótulo, `rotulo_pt` e
`rotulo_en` — no gráfico PT você usa uma como eixo de categorias, no EN a outra.
Mesma tabela, mesmos números.

Nas tabelas `s_`, o cabeçalho de cada coluna é o rótulo em português (é dele que
o Excel tira o nome da série). Para a versão inglesa, aponte o nome da série
para a célula correspondente em `d_<pergunta>[rotulo_en]`.

**`atual` e `anterior` já vêm lado a lado**, com o `delta`. A comparação com o
mês passado não precisa de fórmula.

### Ao montar o gráfico, use a coluna da Tabela

Não um intervalo de células. Clicando no cabeçalho da coluna da Tabela, o
gráfico passa a acompanhar quando a tabela cresce ou encurta — e ela encurta,
porque alternativa aposentada sai. Com intervalo fixo você teria linha em branco
ou dado cortado.

---

## Número publicado não muda

```
ondas até ULTIMA_ONDA_PUBLICADA  ->  valor CONGELADO (o que foi publicado)
ondas novas                     ->  valor CALCULADO do dado bruto
```

**Conferido: 3.574 de 3.574 valores publicados saem idênticos**, cobrindo as 76
ondas de fev/2020 a jul/2026. O `reconciliar.py` mede isso a qualquer momento —
o BLOCO 1 tem que dar 100%.

Por que existe esse regime: a aba `Base` da planilha antiga cobre 76 ondas, mas
a `Raw Data` só 37. São **39 ondas — 3 anos e 5 meses — sem dado bruto atrás**,
que não têm como ser recalculadas.

A coluna `fonte` de `PA Base.xlsx` diz de onde cada número veio, e
`pct_calculado` mostra ao lado o que o recálculo daria.

### O aviso que aparece toda rodada

O `atualizar.py` compara publicado com recalculado na onda corrente e avisa onde
discordam muito. Em jul/2026 são três:

| | Publicado | Recalculado | |
|---|---|---|---|
| `classes_ativos / Tesouro Direto e Renda Fixa` | 57,9% | **84,1%** | +26,2 pp |
| `riscos_bolsa / Outra` | 100,0% | **0,9%** | −99,1 pp |
| `apetite_risco / Mudança de rumo na política econômica` | 15,9% | **58,9%** | +43,0 pp |

As tabelas trazem o **publicado**, porque a onda está congelada. Se quiser os
corrigidos, é decisão de republicação: baixe `ULTIMA_ONDA_PUBLICADA` no
`comum.py` e rode de novo.

Os três são erros da planilha antiga, não do pipeline — o primeiro e o terceiro
são o bug do `;` na contagem de múltipla escolha; o segundo é a linha "Outra",
que publicava 100%.

---

## Como o pipeline conta

Três coisas que o processo antigo errava:

**Múltipla escolha é comparada token a token.** A célula é separada por `;` e
cada pedaço normalizado (sem acento, sem maiúscula, sem `;` sobrando). Sobra ou
falta de separador no fim é indiferente. O processo antigo usava o padrão
`"*" & alternativa & ";*"`, que exige um `;` **depois** da alternativa — quando o
export não punha, a última alternativa marcada por cada respondente não era
contada.

**O denominador é quem respondeu aquela pergunta**, não o total da onda. Em
jul/2026, `interesse_internacional` foi respondida por 105 dos 107.

**Texto livre vai para "Outra", mas com trava.** Resposta digitada à mão cai na
"Outra". Se **mais de 20%** das respostas caírem lá, isso não é texto livre — é
sinal de que o conjunto de alternativas daquela onda não é o do registro. Aí o
script para (onda nova) ou pula o bloco deixando o publicado (onda congelada).
Foi o que aconteceu com `apetite_risco`, redesenhado em 2025.

### Quais alternativas entram no gráfico

"Zero neste mês" e "aposentada" são coisas diferentes, e a regra depende do tipo
de pergunta:

- **Ranking** (`classes_ativos`, `riscos_bolsa`, `setores`,
  `interesse_internacional`, `apetite_risco`) — entram só as que a onda corrente
  usou de fato. É isso que tira as sete alternativas de `riscos_bolsa` que a
  planilha antiga manteve zeradas desde o cutover de abr/2026.
- **Enumeração fixa** (`regiao`, faixas, escala 0-10, `ibovespa_alvo`) — entram
  todas que tiveram resposta nas últimas 12 ondas. Região sem nenhuma resposta
  num mês continua na pizza.

E quando uma alternativa foi renomeada, só a versão atual entra — senão o
gráfico mostraria a mesma barra duas vezes. Na série temporal as duas viram
**uma coluna só**, então a linha não se parte no mês da renomeação.

---

## Quando a pesquisa mudar

Você mexe em **um arquivo**: `registro.csv`. Uma linha por alternativa.

| Coluna | Para que serve |
|---|---|
| `pergunta_id` | Qual pergunta. |
| `alternativa_id` | A identidade. **Nunca muda** — é ela que segura o histórico. |
| `serie_id` | Junta alternativas que são o mesmo conceito com nomes diferentes ao longo do tempo. É o que mantém a série temporal contínua. |
| `rotulo_pt` / `rotulo_en` | Como aparece no Forms e no slide. |
| `aliases` | Nomes **antigos** do mesmo rótulo, separados por `\|`. |
| `valor_num` | Só em `ibovespa_alvo`: o ponto médio da faixa, para a resposta média. |
| `ordem` | Ordem de exibição nas perguntas de enumeração fixa. |

### A tabela de decisão

| O que aconteceu | O que fazer |
|---|---|
| Mudou o texto de uma alternativa | Texto novo em `rotulo_pt`, o antigo vai para `aliases`. |
| Entrou alternativa nova | Uma linha nova. **Só isso** — não tem endereço para acertar. |
| Saiu alternativa | Nada. Ela sai dos gráficos sozinha. |
| Alternativa é o mesmo conceito de outra, com outro nome | Ponha o mesmo `serie_id` nas duas. |
| Mudou o texto da **pergunta** | Acrescente um padrão em `match` no `comum.py`. Não apague os antigos. |
| Ibovespa rolou de 2026 para 2027 | Atualize os 6 `rotulo_pt` de `ibovespa_alvo` e os `valor_num`. |
| Entrou pergunta do mês | Nada. Cai na aba `q_mes`. |
| Entrou pergunta recorrente nova | Um bloco em `BLOCOS` (`comum.py`) e as linhas dela no `registro.csv`. As tabelas `d_` e `s_` aparecem sozinhas; aí você monta o gráfico. |

**Não existe mais regra de "não insira no meio".** As tabelas são geradas
inteiras a cada rodada e os gráficos apontam para colunas de Tabela, então nada
depende de posição.

Se o Python passar a gerar uma tabela nova, rode o `montar_charts.ps1` num
arquivo novo para ver a consulta pronta, ou crie a consulta à mão copiando o
padrão de uma existente.

---

## Mapa dos arquivos

```
pesquisa_assessores\
├── rodar.bat                   ← duplo clique todo mês
├── LEIA-ME.md                  ← este arquivo
├── registro.csv                ← O ARQUIVO QUE VOCÊ EDITA
├── ibovespa.csv                ← uma linha por mês, o fechamento
├── historico_congelado.csv     ← o publicado (gerado 1x)
├── comum.py                    ← caminhos e as perguntas recorrentes
├── atualizar.py                ← o motor
├── reconciliar.py              ← confere contra o publicado
├── congelar.py                 ← roda 1x, na instalação
├── montar_charts.ps1           ← roda 1x, cria a PA Charts.xlsx
├── _logs\                      ← log de cada rodada
├── _saida\                     ← base completa publicada (auditoria)
└── _v1_descartada\             ← tentativas anteriores, arquivadas
```

Na rede:

```
Pesquisa assessores\
├── PA Charts.xlsx           ← SUA planilha. os gráficos vivem aqui.
├── PA Principal.xlsx        ← fonte do histórico. não apague (ver abaixo).
├── input_forms\             ← o export do Forms
└── bases\
    ├── PA Base.xlsx         ← a base geral
    ├── PA Charts Data.xlsx  ← as mesmas tabelas, para folhear
    └── charts\*.csv         ← a fonte do Power Query. não apague.
```

**Sobre a PA Principal:** ela não é lida por nada no fluxo mensal. Mas é dela
que saiu o `historico_congelado.csv`, e sem ela não dá para regerar esse arquivo
— são as 39 ondas que não têm dado bruto. Guarde.

---

## Instalação

Precisa de Python 3.9+ com `openpyxl` (já instalado), e Excel (para o
`montar_charts.ps1`).

Uma vez só, nesta ordem:

```bash
python congelar.py
```

Lê a `PA Principal.xlsx` e grava `registro.csv`, `historico_congelado.csv` e
`ibovespa.csv`. Confira os avisos.

```bash
python atualizar.py --bootstrap
```

Traz as 37 ondas de respostas cruas da `Raw Data` para a base geral.

```bash
python reconciliar.py
```

O BLOCO 1 tem que dar 100%.

```bash
powershell -ExecutionPolicy Bypass -File montar_charts.ps1
```

Cria a `PA Charts.xlsx` com as 27 consultas prontas. Leva uns **3 a 4 minutos** —
é o Excel montando consulta por consulta, e roda só esta vez. Ele **se recusa a
rodar se o arquivo já existir**, então não tem como apagar seus gráficos por
acidente.

Depois disso é sua vez: monte os gráficos sobre as tabelas. E daí em diante, todo
mês, só o `rodar.bat` + Atualizar Tudo.

### Testar sem a rede

```bash
set PA_REDE=C:\temp\rede_teste
```
