# Migração da PA Principal.xlsx — passo a passo

Esta é a única parte que precisa ser feita à mão, uma vez só.

## A restrição que manda em tudo

Os dois PPTs têm **11 links OLE cada (22 no total)**, e cada link aponta para
uma coordenada exata:

```
\\xpdocs\...\PA Principal.xlsx ! Charts ! [PA Principal.xlsx]Charts Gráfico 1-1
                ↑ caminho          ↑ aba        ↑ nome do objeto de gráfico
```

Isso significa três coisas, e elas não são negociáveis:

1. **A PA Principal.xlsx continua sendo o arquivo linkado.** Planilha nova = 22 links mortos.
2. **A aba `Charts` não muda de nome** e os objetos de gráfico dela não são apagados nem recriados.
3. Tudo que a gente muda fica **por baixo** da aba `Charts`: na aba `Base`.

Se um link quebrar, o conserto é: no PPT, botão direito no gráfico >
*Editar Links* > apontar de novo. Dá para fazer, mas é chato — o desenho
abaixo evita chegar nesse ponto.

---

## Antes de começar

Rode os três comandos da seção "Instalação" do `LEIA-ME.md`, nesta ordem:
congelar → bootstrap → reconciliar. O `reconciliar.py` tem que fechar o
bloco 1 em 100% antes de você encostar na planilha.

Isso garante o que importa: **nenhum número já publicado muda.** As ondas
históricas saem congeladas com o valor que foi ao ar; só as ondas novas são
calculadas do bruto.

---

## Etapa 1 — religar a Base ao agregado

1. Na PA Principal, crie a consulta `PastaBases` (bloco [1] de `consultas.m`)
   e carregue como **Apenas Criar Conexão**.

2. Crie a consulta `agregado` (bloco [2]) e carregue **numa aba nova** chamada
   `agregado`. Confirme o nome da tabela em *Design da Tabela > Nome da Tabela*.

3. Cole a coluna `chave` do arquivo `_saida/chaves_base.csv` (gerado pelo
   `congelar_historico.py`) na **coluna CC** da aba Base, a partir da linha 1.

   `CC` é a primeira coluna livre à direita — a Base usa até `CB` hoje. Colar
   ali não insere nada e **nenhuma referência existente se desloca**: nem as da
   aba `Charts`, nem as da `Gráfico capa`.

4. Numa coluna de onda qualquer (ex.: `CB`, jul/26), na primeira linha de
   alternativa (ex.: `CB8`), troque a fórmula por:

```
=IF(COUNTIFS(agregado[chave],$CC8,agregado[onda],CB$1)=0,"",SUMIFS(agregado[pct],agregado[chave],$CC8,agregado[onda],CB$1))
```

   Arraste para baixo em todas as linhas que têm chave em `CC`, e para a
   esquerda até a primeira coluna de onda que quiser religar.

   Use a coluna `pct` — é ela que respeita o congelamento. A `pct_calculado`
   é só auditoria; não aponte gráfico para ela.

   Por que `SUMIFS` e não `INDEX/MATCH`: o par (chave, onda) é único, então o
   `SUMIFS` devolve o valor exato — e ele não depende de casar cabeçalho de
   coluna, que quebraria com nome de mês ou Excel em outro idioma.

5. Linhas com `CC` vazio são blocos de perguntas do mês antigas. Deixe as
   fórmulas velhas ali ou apague o bloco. Não atrapalham.

6. **Confira antes de seguir:** os valores das colunas de onda antigas têm que
   continuar exatamente como estavam. Se algum mudou, pare — ou a chave em `CC`
   está desalinhada, ou o congelamento não foi gerado.

7. A partir daqui, **mês novo não exige arrastar nada**: você só põe o código
   da onda (ex.: `202608`) na linha 1 de uma coluna nova e copia a fórmula da
   coluna anterior. O denominador correto já vem do `agregado`.

---

## Etapa 2 — os gráficos da pergunta do mês

1. Crie a consulta `q_mes` (bloco [7]) e carregue numa aba nova `q_mes`.

2. Os slots têm endereço fixo. Com `linhas_por_slot: 12` (o padrão):

   | Slot | Título | Rótulos PT | Rótulos EN | Valores |
   |------|--------|-----------|-----------|---------|
   | 1 | `q_mes!$A$2` | `q_mes!$A$4:$A$15` | `q_mes!$B$4:$B$15` | `q_mes!$C$4:$C$15` |
   | 2 | `q_mes!$A$17` | `q_mes!$A$19:$A$30` | `q_mes!$B$19:$B$30` | `q_mes!$C$19:$C$30` |
   | 3 | `q_mes!$A$32` | `q_mes!$A$34:$A$45` | `q_mes!$B$34:$B$45` | `q_mes!$C$34:$C$45` |
   | 4 | `q_mes!$A$47` | `q_mes!$A$49:$A$60` | `q_mes!$B$49:$B$60` | `q_mes!$C$49:$C$60` |

   (fórmula geral: o slot *s* começa na linha `2 + (s-1) × 15`)

3. Monte **uma vez** um gráfico por slot na aba `Charts`, apontando para esses
   endereços. Cole no PPT como link, como você já faz.

4. Todo mês o gráfico se atualiza sozinho. Linhas sobrando ficam em branco:
   configure em *Selecionar Dados > Células Ocultas e Vazias > Mostrar como:
   Vazio*.

**Não ordene, filtre nem insira linhas na aba `q_mes`** — os gráficos apontam
para endereços absolutos.

Continua manual: a **tradução** da pergunta do mês. Ou você digita na coluna B,
ou declara em `perguntas_mes` no `config/perguntas.yaml` e ela passa a sair
pronta.

---

## Etapa 3 — opcionais, quando quiser

- **`serie`** (bloco [6]): janela móvel para os gráficos de linha, incluindo a
  aba `Gráfico capa`. Já vem com a data pronta.
- **`meta`** (bloco [8]): título do mês em PT e EN, nº de respostas. Use nos
  títulos dos slides em vez de digitar o mês na mão.
- **`raw_norm`** (bloco [4]): a Raw Data com cabeçalho canônico e alternativas
  já normalizadas. **Não use para alimentar a Base** — as fórmulas COUNTIFS
  recalculariam o histórico e mudariam número publicado. Serve para conferência
  e para recorte ad-hoc.
- **`respostas`** (bloco [10]): o grão respondente a respondente, ~124 mil
  linhas. Para cortes por região, por perfil de alocação, etc.

---

## Atualização mensal, depois de tudo montado

1. Exporte o Forms para `input_forms\`
2. Duplo clique em `rodar.bat`
3. Abra a PA Principal > **Dados > Atualizar Tudo**
4. Abra os PPTs > **Atualizar Links**

Se o passo 2 parar com erro, é porque entrou alternativa nova na pesquisa. O
log diz qual é e o que fazer. Nada é gravado até você resolver.
