# Uma consulta por arquivo

Cada `.m` desta pasta é **uma consulta inteira**, pronta para colar sem
escolher pedaço. Cole o arquivo todo, comentários e tudo — o Power
Query entende `//` como comentário.

## O caminho de cliques

Para cada arquivo, na ordem numérica:

1. **Dados › Obter Dados › De Outras Fontes › Consulta em Branco**
2. Abre o Editor do Power Query. Vá em **Página Inicial › Editor
   Avançado**
3. **Apague tudo** que estiver lá (`Ctrl+A`, `Delete`)
4. Cole o conteúdo do arquivo
5. **Concluído**
6. No painel da direita, em *Propriedades › Nome*, escreva o nome da
   consulta — exatamente como está no arquivo, sem o número:
   `PastaBases`, `paineis`, `tendencias`, `q_mes`, `meta`, `layout`
7. **Página Inicial › Fechar e Carregar Em…** e escolha conforme a
   tabela abaixo

## O modo de carga muda entre elas

| Arquivo | Nome da consulta | Fechar e Carregar Em… |
|---|---|---|
| `01_PastaBases.m` | `PastaBases` | **Apenas Criar Conexão** |
| `02_paineis.m` | `paineis` | **Tabela** › Nova planilha |
| `03_tendencias.m` | `tendencias` | **Tabela** › Nova planilha |
| `04_q_mes.m` | `q_mes` | **Tabela** › Nova planilha |
| `05_meta.m` | `meta` | **Tabela** › Nova planilha |
| `06_layout.m` | `layout` | **Tabela** › Nova planilha |

Só a primeira é conexão. Consulta em modo conexão não põe nada em
célula — o dado fica no cache interno. E os gráficos apontam para
**endereço de célula** (`paineis!$A$5:$A$22`); se estas cinco forem
conexão, não haverá o que ler.

A `PastaBases` pode ser conexão porque não é dado: é um texto só, o
caminho da pasta, que as outras cinco usam como parâmetro. **Crie ela
primeiro** — as outras quebram sem ela.

Depois de renomear a aba criada, confira que o nome ficou igual ao da
consulta. O Excel às vezes cria como `paineis (2)` se já existir algo
parecido, e aí os endereços do `layout` não batem.

---

## Dois jeitos de errar (e o sintoma de cada um)

### A tabela mostra o código em vez dos números

Você importou o `.m` **como arquivo de dados** — algo como
*Obter Dados › De Arquivo › De Texto*. O Power Query leu o arquivo como
texto e transformou cada linha do código numa linha da tabela. Por isso
aparecem `let`, `Arquivo = ...` e os comentários.

O `.m` não é fonte de dados. Ele é o código da consulta, e só entra
pelo **Editor Avançado**.

Apague essa consulta e refaça pelo caminho de cliques acima.

### Erro de sintaxe ao colar

Provavelmente você colou o `consultas.m` da pasta de cima, que tem os
**10 blocos juntos**. Cada bloco é uma consulta separada; emendados não
formam M válido.

Use os arquivos desta pasta — cada um já é uma consulta só. O
`consultas.m` continua lá como referência de leitura.

---

## As quatro opcionais

`07` a `10` não alimentam gráfico nenhum. São para análise ad-hoc:
histórico completo, série longa, médias e o grão respondente a
respondente. Carregue só se for usar.
