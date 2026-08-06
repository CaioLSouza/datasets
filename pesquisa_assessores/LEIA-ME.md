# Pesquisa de Assessores XP — processo novo

## O fluxo, em 4 passos

```
1. Exportar o Forms  ──►  input_forms\
2. Duplo clique em   ──►  rodar.bat
3. PA Report.xlsx    ──►  Dados > Atualizar Tudo
4. Os dois PPTs      ──►  Atualizar Links
```

O passo 2 é o único que é novo. Ele substitui "colar na Raw Data, arrastar as
fórmulas para a direita, conferir se apareceu pergunta nova".

A `PA Report.xlsx` é uma planilha nova, montada uma vez (ver
`powerquery/INSTRUCOES.md`). A **PA Principal.xlsx antiga sai de cena** — vira
arquivo, não é mais tocada. Os motivos estão logo abaixo.

---

## O que o `rodar.bat` faz

Lê o export do Forms e gera dois arquivos na pasta `bases\`:

| Arquivo | Abas | Tamanho |
|---|---|---|
| `PA Base Historica.xlsx` | `agregado` (todo o histórico), `matriz` (chave × onda), `respostas` (grão respondente, ~124 mil linhas), `raw_norm` (a Raw Data limpa), `medias`, `dicionario` | ~6 MB |
| `PA Base Mes Atual.xlsx` | **`paineis`**, **`tendencias`**, `q_mes`, **`layout`**, `meta`, `report`, `serie`, `medias` | ~150 KB |

E um log em `_logs\rodada_AAAAMM.txt` dizendo exatamente o que reconheceu.

**Se entrar alternativa nova na pesquisa, ele para e não grava nada.** O log
diz qual é e o que fazer. A base nunca fica meio atualizada.

### As três abas que sustentam os gráficos

- **`paineis`** — um retângulo de endereço fixo por pergunta, com pct, pct do
  mês anterior e delta. 12 blocos.
- **`tendencias`** — as séries temporais em colunas fixas, uma por item de
  `series_do_report` no config. Alimenta os gráficos de linha e a capa.
- **`q_mes`** — 4 slots de endereço fixo para a pergunta que muda todo mês.

Endereço fixo é o ponto: **o gráfico é montado uma vez e nunca mais precisa ser
refeito** — nem quando a pergunta ganha alternativa, nem quando ela falta num
mês. Um mês sem `selic_alvo`, por exemplo, deixa o bloco dela vazio no lugar de
sempre; nada abaixo se desloca.

A aba **`layout`** lista os 21 endereços prontos para copiar, já com o
deslocamento do Power Query aplicado. Você nunca conta linha.

---

## Por que a PA Principal foi aposentada

O diagnóstico dela, medido nos dados — é o que justificou recomeçar em vez de
remendar:

**A Raw Data tinha 79 colunas, 55 delas ~98% vazias.** Cada pergunta do mês
virava uma coluna nova para sempre. Em 3 anos foram 37 ondas e 6.735 respostas
espalhadas numa matriz quase toda em branco.

**A ligação pergunta ↔ coluna era por texto exato.** A fórmula fazia
`MATCH(texto_da_pergunta; 'Raw Data'!$1:$1; 0)`. Qualquer letra diferente no
Forms criava coluna nova. Aconteceu: *"os maiores risco"* virou *"o maiores
riscos"*; *"apetite por risco"* virou *"apetite a risco"*.

**Em abr/2026 o conjunto de alternativas foi reescrito.** Cinco alternativas de
`riscos_bolsa` mudaram de rótulo mantendo o conceito — *"Riscos geopolíticos"*
→ *"Riscos geopolíticos/ Guerra"*, *"Instabilidade política"* →
*"Instabilidade política/ Eleições"*, e mais três. As janelas de onda são
disjuntas (uma até 202603, a outra a partir de 202604), o que prova que é
renomeação. Na Base cada uma virou linha nova, então **a série histórica dessas
alternativas zera em abr/2026** e recomeça do lado.

**No `apetite_risco` é pior:** "corte de juros nos EUA" já teve *quatro*
rótulos diferentes, todos disjuntos. A série está picotada em quatro pedaços.

**O denominador era digitado à mão** na linha 2 da Base, e era o total da onda,
não quem respondeu aquela pergunta. Em fev/2026, três perguntas foram
respondidas por 100 dos 107 — os percentuais saíram divididos por 107.

### O bug que estava mudando números publicados

A Base conta múltipla escolha com o padrão `"*" & alternativa & ";*"`. Esse
padrão exige um `;` **depois** da alternativa. Quando o export do Forms não põe
`;` no fim da célula, a **última alternativa marcada por cada respondente não é
contada**.

Dois casos conferidos na unha em jul/2026 (o valor publicado bate exatamente
com a contagem truncada):

| Alternativa | Publicado | Real | Erro |
|---|---|---|---|
| Tesouro Direto e Renda Fixa | 57,9% (62/107) | **84,1%** (90/107) | −26 pp |
| Mudança de rumo na política econômica | 15,9% (17/107) | **58,9%** (63/107) | −43 pp |

A condição que dispara o problema — célula com `;` no meio mas sem `;` no fim —
aparece em **84,6% das células de múltipla escolha em jul/2026** e em **40,5%
em mar/2026**, com incidência menor em out/2025, fev/2026 e mai/2026. Nas
demais ondas não ocorre.

O pipeline não tem como errar isso: ele separa por `;` e compara token a token,
então sobra ou falta de separador é indiferente.

### O custo de recomeçar

Um só, e é uma vez: **os 22 links OLE dos PPTs precisam ser refeitos** — 11 por
PPT. A `powerquery/INSTRUCOES.md` tem um atalho que costuma resolver os 11 de
uma vez, porque copiamos a aba `Charts` inteira e os nomes dos objetos vão
junto.

O visual dos slides não muda: os gráficos são os mesmos, só passam a olhar
para outro lugar.

---

## Número publicado não muda

Nada do que já foi ao ar é restatement. O pipeline trabalha em dois regimes:

```
ondas até `ultima_onda_publicada`  ->  valor CONGELADO (o que foi publicado)
ondas novas                        ->  valor CALCULADO do dado bruto
```

O congelamento vem de `config/valores_publicados.csv`, gerado uma vez pelo
`src/congelar_historico.py` lendo a própria aba Base. Cada linha do `agregado`
carrega a coluna `fonte` dizendo de onde o número veio.

Conferido: **2.893 de 2.894 valores publicados saem idênticos.** O `reconciliar.py`
mede isso a qualquer momento.

O único caso que sobra precisa da sua decisão: em jul/2026 a Base tem **duas
linhas** para *"Melhora na recuperação econômica global"*, publicando 12,15% e
21,50%. A de cima tem um `;` colado no rótulo e por isso subconta; a de baixo é
a correção. O congelador fica com a de baixo e avisa. Se você preferir a outra,
é só editar o valor em `config/valores_publicados.csv`.

### E os merges de alias, não mexem em número?

Não. Como as janelas são disjuntas — *"Riscos geopolíticos"* até 202603 e
*"Riscos geopolíticos/ Guerra"* de 202604 em diante — juntar as duas na mesma
linha só emenda a série. Cada mês continua com o valor que foi publicado. Você
ganha a série contínua sem restatement.

### E as correções, quando entram?

Só nas ondas novas. O recálculo do bruto fica visível na coluna
`pct_calculado` do `agregado`, lado a lado com o publicado, para auditoria —
mas não alimenta gráfico nenhum enquanto a onda estiver congelada. Se um dia
você decidir restatear alguma coisa, é uma linha no `config.yaml`.

---

## Quando a pesquisa mudar

Você mexe em **um arquivo só**: `config/perguntas.yaml`. Ele tem as instruções
por dentro. Resumo:

| O que aconteceu | O que fazer |
|---|---|
| Mudou o texto da pergunta | Acrescenta um padrão em `match:`. Não apaga os antigos. |
| Alternativa mudou de rótulo | Move o rótulo velho para `aliases:`, põe o novo em `pt:`. A série continua inteira. |
| Entrou alternativa nova | Adiciona em `opcoes:` com `id`, `pt` e `en`. |
| Saiu alternativa | Não faz nada. Ela só para de aparecer. |
| **Ibovespa rolou de 2026 para 2027** | **Nada.** O regex já pega qualquer ano. |
| Entrou pergunta do mês | **Nada.** Vira slot automaticamente. |
| Entrou pergunta **recorrente** nova | Adiciona com o **maior `ordem`** da lista. Ver abaixo. |

O `id` de uma pergunta ou alternativa **nunca muda** depois de criado — é ele
que segura o histórico.

### A única regra que protege os gráficos

Os blocos da aba `paineis` saem na ordem do campo `ordem` do registro. Uma
pergunta recorrente nova com `ordem` menor que as existentes **empurra os
blocos de baixo** e quebra os gráficos já montados.

Então: pergunta nova recebe `ordem` maior que todas. Hoje a maior é 120
(`selic_alvo`), então a próxima é 130. O bloco entra no fim e nada se move.

Mesma lógica em `series_do_report` (config.yaml): acrescente sempre **no fim da
lista**, nunca no meio.

Se um painel estourar as `linhas_por_painel` (hoje 18, e a maior pergunta usa
15), o pipeline **para de rodar e diz quais alternativas ficariam de fora**. Ele
não trunca calado.

### O Ibovespa rolando de ano

Na **base histórica**, cada safra é uma série separada (`safra=2025`,
`safra=2026`…) — você nunca perde o que foi perguntado na época.
Na **base do mês**, só entra a safra corrente, no mesmo endereço de sempre —
o gráfico do report não precisa ser refeito.

As faixas ("Entre 190 e 200 mil pontos") são lidas por um parser numérico, não
cadastradas uma a uma. Quando em abr/2026 as faixas de 10 mil viraram faixas de
20 mil, nada precisou ser configurado.

### A pergunta do mês

Toda coluna do Forms que não casa com nenhuma pergunta conhecida vira slot
automaticamente. Os slots têm **endereço fixo** na aba `q_mes` — o gráfico é
montado uma vez e todo mês só muda o conteúdo. Endereços na
`powerquery/INSTRUCOES.md`.

Continua manual: a tradução da pergunta do mês para o inglês.

---

## Mapa dos arquivos

```
pesquisa_assessores\
├── rodar.bat                   ← duplo clique todo mês
├── LEIA-ME.md                  ← este arquivo
├── config\
│   ├── config.yaml             ← caminhos de rede e parâmetros
│   ├── perguntas.yaml          ← O ARQUIVO QUE VOCÊ EDITA
│   └── valores_publicados.csv  ← o histórico congelado (gerado 1x)
├── src\
│   ├── pipeline.py             ← o motor
│   ├── congelar_historico.py   ← congela o publicado + gera as chaves da Base
│   └── reconciliar.py          ← confere a saída contra o publicado
├── powerquery\
│   ├── consultas.m             ← as consultas da PA Report.xlsx
│   └── INSTRUCOES.md           ← como montar a planilha nova, 1 vez
├── _dados\respostas.csv        ← fonte da verdade (local, append-only)
├── _logs\                      ← log de cada rodada
└── _saida\chaves_base.csv      ← o mapa da Base antiga (auditoria)
```

Na rede, ao lado da PA Principal antiga, aparecem duas coisas:

```
Pesquisa assessores\
├── PA Principal.xlsx        ← ARQUIVO. não é mais tocada.
├── PA Report.xlsx           ← a planilha nova, linkada pelos PPTs
├── input_forms\             ← você larga o export do Forms aqui
└── bases\                   ← o pipeline escreve aqui
    ├── PA Base Historica.xlsx
    └── PA Base Mes Atual.xlsx
```

---

## Instalação

Precisa de Python 3.9+ com dois pacotes:

```bash
pip install openpyxl pyyaml
```

Se `pip` estiver bloqueado pela política de TI, peça a instalação desses dois —
são pacotes padrão, sem dependência externa e sem acesso a rede.

Depois, dois comandos, uma vez só, nesta ordem.

**1. Congelar o que já foi publicado:**

```bash
python src\congelar_historico.py "\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores\PA Principal.xlsx"
```

Lê a aba Base e grava `config/valores_publicados.csv`. Confira o aviso de
conflito que ele imprime no fim.

**2. Importar as 37 ondas de respostas:**

```bash
python src\pipeline.py --bootstrap "\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores\PA Principal.xlsx"
```

**3. Conferir que nada mudou:**

```bash
python src\reconciliar.py "\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores\PA Principal.xlsx"
```

O bloco 1 tem que dar 100%. É ele que prova que nenhum número publicado mudou.

**4. Montar a `PA Report.xlsx`:**

Siga a `powerquery/INSTRUCOES.md`. São três etapas: criar as 6 consultas, trazer
os gráficos da PA Principal com o visual intacto, e refazer os 22 links dos
PPTs.

Depois disso é só o `rodar.bat` todo mês.

---

## Sugestões que ficaram de fora, para depois

Coisas que valem, mas que não fazem parte do que foi pedido:

1. **Fechamento do Ibovespa na capa.** Hoje é digitado à mão na coluna E da aba
   `Gráfico capa`. Já existe um campo `ibovespa_fechamento` em `config.yaml`
   para você preencher uma linha por mês — de lá ele já flui para a base do
   mês. Automatizar a busca do fechamento depende de acesso a alguma fonte de
   dados da casa.

2. **Rodar sozinho.** O `rodar.bat` pode ir para o Agendador de Tarefas
   apontando para a pasta `input_forms`. Ele devolve código de erro quando algo
   precisa de atenção, então dá para receber aviso.

3. **Confirmar a linha duplicada de `apetite_risco`.** A Base antiga tem duas
   linhas para *"Melhora na recuperação econômica global"*, publicando 12,15% e
   21,50% em jul/2026. O congelador ficou com 21,50% (a linha de baixo, sem o
   `;` colado no rótulo) e avisa toda vez que roda. Se você confirmar que é essa
   mesmo, o aviso pode ser ignorado para sempre — só não vale esquecer dele.

4. **Padronizar as alternativas no próprio Forms.** Metade dos aliases existe
   porque o texto da alternativa foi reescrito sem necessidade. Congelar o
   texto no Forms é mais barato que cadastrar alias depois.

5. **Um dia, restatear o histórico.** Se em algum momento fizer sentido
   republicar a série corrigida (sem o bug do `;`, com denominador certo), é
   uma linha: baixe `ultima_onda_publicada` no `config.yaml` e rode de novo. A
   coluna `pct_calculado` já mostra hoje o que sairia. Enquanto isso não for
   uma decisão sua, nada muda.

6. **Arquivar a PA Principal.** Ela ainda é necessária: é dela que sai o
   congelamento (`congelar_historico.py`) e dela que vieram os gráficos. Depois
   que a PA Report estiver rodando há uns meses, mova-a para uma subpasta
   `arquivo\` — mas **não apague**: sem ela não dá para regerar o
   `valores_publicados.csv` do zero.
