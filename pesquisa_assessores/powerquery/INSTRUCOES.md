# Montar a PA Report.xlsx do zero

Isto se faz **uma vez**. Depois, todo mês é só `rodar.bat` + Atualizar.

A PA Principal.xlsx antiga não é mais tocada nem apagada — ela vira
arquivo. Guarde-a: é dela que sai o congelamento do histórico, e é dela
que vamos copiar os gráficos.

---

## Por que jogar a PA Principal fora

Ela carrega quatro problemas que nenhum remendo resolve:

- **O COUNTIF perde a última opção marcada.** O padrão é
  `"*" & alternativa & ";*"`, que exige `;` *depois* da alternativa.
  Quando o Forms não põe `;` no fim da célula, a última opção de cada
  respondente não entra na conta. Atinge 84,6% das células de múltipla
  escolha em jul/2026.
- **11 blocos mortos** de perguntas do mês antigas, ocupando linhas.
- **Uma linha duplicada** em `apetite_risco` publicando 12,15% e 21,50%
  para a mesma alternativa, ao mesmo tempo.
- **As fórmulas acham a pergunta por texto exato** (`MATCH` no cabeçalho
  da Raw Data). Qualquer letra diferente no Forms cria coluna nova.

O preço de recomeçar é **refazer os 22 links OLE** — 11 por PPT. A
Etapa 3 abaixo mostra um atalho que costuma resolver os 11 de uma vez.

---

## Etapa 0 — gerar as bases

Se ainda não fez, rode os três comandos da seção "Instalação" do
`LEIA-ME.md`: congelar → bootstrap → reconciliar.

O `reconciliar.py` tem que fechar o bloco 1 em ~100% (hoje: **3.717 de
3.718**). Isso é o que garante que nenhum número já publicado mudou de
valor na virada.

O histórico cobre **76 ondas, de fev/2020 a jul/2026** — a aba Base
alcança bem mais para trás que a Raw Data, que só começa em jul/2023.

Ao fim você tem, em `\\xpdocs\...\Pesquisa assessores\bases\`:

- `PA Base Historica.xlsx`
- `PA Base Mes Atual.xlsx`  ← é dela que o report se alimenta

---

## Etapa 1 — a planilha nova e as consultas

1. Crie uma pasta de trabalho nova em
   `\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores\`
   chamada **`PA Report.xlsx`**.

   O nome importa: é ele que vai dentro dos 22 links dos PPTs. Depois
   de linkado, **não renomeie e não mude de pasta**.

2. Crie a consulta `PastaBases` (bloco [1] de `consultas.m`) e carregue
   como **Apenas Criar Conexão**.

3. Crie as cinco seguintes — `paineis`, `tendencias`, `q_mes`, `meta`,
   `layout` — e carregue cada uma **numa aba nova de mesmo nome**.

   Nenhuma delas promove cabeçalho, e isso é de propósito: as três
   primeiras têm endereço fixo, e promover cabeçalho deslocaria tudo
   em uma linha. Está explicado no topo do `consultas.m`.

4. Abra a aba `layout`. Ela lista, para cada gráfico, exatamente qual
   intervalo apontar — com o deslocamento do Power Query já aplicado.
   Você não precisa contar linha nenhuma.

   São 21 fontes de gráfico: 12 painéis de pergunta, 5 séries
   temporais e 4 slots de pergunta do mês.

---

## Etapa 2 — trazer os gráficos, com o visual intacto

A ideia é **não redesenhar nada**. Os gráficos da PA Principal já têm
a formatação certa; só precisam olhar para outro lugar.

1. Abra a PA Principal.xlsx e a PA Report.xlsx lado a lado.

2. Na PA Principal, botão direito na aba `Charts` > **Mover ou Copiar**
   > para a PA Report.xlsx > marque **Criar uma cópia**.

   Isso traz os gráficos com toda a formatação **e preserva os nomes
   dos objetos** ("Gráfico 1-1" etc.). Guardar esses nomes é o que
   torna a Etapa 3 fácil — não renomeie nenhum deles.

3. Faça o mesmo com a aba `Gráfico capa`.

4. Os gráficos vão chegar apontando para a PA Principal antiga. Um por
   um, clique no gráfico > **Selecionar Dados** > e troque cada série
   para o endereço correspondente da aba `layout`:

   | O que a aba `layout` dá | Onde entra em Selecionar Dados |
   |---|---|
   | `rótulos_pt` (ou `rótulos_en`, no PPT em inglês) | Rótulos do Eixo Horizontal |
   | `valores` | Valores da série |
   | `título` | Nome da série |
   | `delta` | segunda série, se o gráfico mostra variação |

5. Nos gráficos de linha, o eixo horizontal é sempre
   `tendencias!$A$5:$A$34` (as datas) — está na coluna `rótulos_pt` das
   linhas de tipo "tendência".

6. Nos quatro slots de pergunta do mês, marque
   **Selecionar Dados > Células Ocultas e Vazias > Mostrar como: Vazio**.
   Assim as linhas não usadas somem do gráfico em vez de virar zero.

7. Apague da PA Report qualquer aba que tenha vindo junto e não seja
   usada. As abas `paineis`, `tendencias`, `q_mes`, `meta` e `layout`
   são as únicas que o Power Query alimenta.

---

## Etapa 3 — os 22 links dos PPTs

Como você copiou a aba `Charts` inteira, os nomes dos objetos são os
mesmos de antes. Só o arquivo mudou. Então tente primeiro o caminho
curto:

1. No PPT, **Arquivo > Informações > Editar Links para Arquivos**
   (ou botão direito num gráfico > *Objeto Gráfico Vinculado* >
   *Links*).

2. Selecione um link > **Alterar Fonte** > aponte para a
   `PA Report.xlsx`.

3. Veja se os outros 10 links do mesmo PPT passaram a apontar para o
   arquivo novo sozinhos. O PowerPoint costuma reapontar todos os
   links que dividiam a mesma origem.

4. Se não passaram, repita o passo 2 para cada um. São 11 por PPT.

5. **Atualizar Links** e confira slide a slide.

Se algum gráfico vier quebrado, o caminho longo é: apagar o objeto no
slide, copiar o gráfico da PA Report e colar com
**Colar Especial > Colar Vínculo**. Isso refaz o link do zero, mas você
perde o posicionamento — anote onde estava antes de apagar.

> Não testei esta etapa: não tenho Excel nem PowerPoint aqui. Os passos
> vêm do formato dos links que li dentro dos dois arquivos. Faça numa
> **cópia** dos PPTs antes de mexer nos originais.

---

## A rotina, depois que estiver montado

1. Exporte o Forms para `input_forms\`
2. Duplo clique em `rodar.bat`
3. Abra a PA Report.xlsx > **Dados > Atualizar Tudo**
4. Abra os PPTs > **Atualizar Links**

Se o passo 2 parar com erro, é porque entrou alternativa nova na
pesquisa. O log diz qual é e o que fazer. Nada é gravado até você
resolver.

---

## O que quebra os endereços (e como não quebrar)

Os gráficos apontam para intervalos absolutos. Três coisas os
deslocam — todas evitáveis:

| Mudança | Efeito | Como fazer sem quebrar |
|---|---|---|
| Pergunta recorrente nova em `perguntas.yaml` | Se o `ordem` for menor que o de alguma existente, empurra os blocos de baixo | Dê a ela o **maior `ordem`** da lista. O bloco entra no fim e nada se move. |
| `linhas_por_painel` no config | Desloca **todos** os painéis | Só mexa se um painel estourar — e aí refaça os endereços pela aba `layout`. |
| Série nova em `series_do_report` | Se inserida no meio, desloca as colunas seguintes | Acrescente sempre **no fim da lista**. |

Se um painel estourar, o pipeline avisa alto na hora de rodar, dizendo
quais alternativas ficaram de fora. Ele não trunca em silêncio.

Ordenar, filtrar ou inserir linha nas abas `paineis`, `tendencias` e
`q_mes` também desloca tudo. Não mexa nelas — são alimentadas pelo
Power Query.
