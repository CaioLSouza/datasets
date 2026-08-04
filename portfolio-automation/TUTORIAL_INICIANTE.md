# Tutorial para iniciantes

Este tutorial ensina a executar a automação no notebook da empresa. Não é
necessário saber programar.

## 1. O que o projeto faz

O processo possui dois comandos:

1. `portfolio_automation.py` lê as bases corporativas, calcula as carteiras,
   gera os Excel, atualiza as Lâminas Comerciais e cria a Prestação de Contas
   de Top Ações, Top Dividendos e Top Small Caps.
2. `email_generator.py` usa esses Excel para criar `email_carteiras.msg`.

O projeto mantém a **Lâmina Comercial** e a **Prestação de Contas**, mas não
gera a lâmina completa auxiliar (`lamina_dados_*.xlsx`). O e-mail não é enviado
automaticamente: ele é salvo para revisão.

## 2. Antes de começar

Confirme:

- Windows conectado à rede da empresa ou VPN;
- acesso às pastas `\\xpdocs`;
- Microsoft Outlook instalado e configurado;
- Microsoft PowerPoint instalado;
- Python instalado;
- Excel e PowerPoint de saída fechados.

Abra o PowerShell e verifique o Python:

```powershell
py --version
```

Se aparecer `Python 3.x`, continue. Caso contrário, use o catálogo de software
da empresa ou peça ajuda ao suporte de TI.

## 3. Baixar e extrair o projeto

Se recebeu `xp-portfolio-automation.zip` por e-mail:

1. Baixe o anexo.
2. Clique com o botão direito e escolha **Extrair tudo**.
3. Coloque a pasta em um local fácil, como `Documents`.
4. Abra a pasta extraída.

Se estiver usando Git:

```powershell
git clone <URL_DO_REPOSITORIO>
cd xp-portfolio-automation
```

## 4. Abrir o PowerShell na pasta correta

No Explorador de Arquivos:

1. Abra a pasta que contém `portfolio_automation.py`.
2. Clique na barra de endereço.
3. Digite `powershell` e pressione Enter.

## 5. Preparar o ambiente

Crie um ambiente isolado:

```powershell
py -m venv .venv
```

Instale as dependências:

```powershell
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Isso normalmente é necessário apenas na primeira utilização.

## 6. Conferir os templates dos PowerPoint

Abra `xp_carteiras\monthly_config.py` e confira os nomes em
`COMMERCIAL_TEMPLATE_FILES`. Eles devem ser idênticos aos arquivos existentes
na pasta de templates, incluindo mês, ano, acentos, espaços e `.pptx`.

Altere somente o nome à direita dos dois-pontos. Não altere o nome da carteira.

No mesmo arquivo, confira `ACCOUNTABILITY_TEMPLATE_FILES`. O projeto já inclui
uma cópia de `Prestação de Contas - Top Ações - Julho 2026.pptx` na pasta
`templates`. Se houver um arquivo com o mesmo nome na pasta corporativa de
templates, o arquivo corporativo terá prioridade.

## 7. Executar os cálculos, Excel e PowerPoint

Confirme a VPN e feche arquivos de saída que estejam abertos. Execute:

```powershell
.\.venv\Scripts\python.exe portfolio_automation.py
```

O processo pode levar alguns minutos. Por padrão, os resultados são gravados
em:

```text
\\xpdocs\Research\Equities\Estrategia\Carteiras\Carteiras de Ações XP\output
```

## 7. Arquivos esperados

Os nomes externos permanecem iguais aos do processo original.

| Grupo | Exemplos |
| --- | --- |
| Curvas base 100 | `top_acoes_base_100.xlsx`, `top_dividendos_base_100.xlsx`, `top_small_caps_base_100.xlsx`, `esg_base_100.xlsx` |
| Métricas | `portfolio_metrics_top_acoes.xlsx` e equivalentes |
| Performance | `tab_performance_top_acoes.xlsx`, `tab_performance_top_dividendos.xlsx`, `tab_performance_top_small_caps.xlsx`, `tab_performance_esg.xlsx` |
| Componentes | `componentes_top_acoes_atual.xlsx`, `componentes_top_acoes_ultimo_rebal.xlsx`, `componentes_top_acoes_comp_mes_passado_mtd_atual.xlsx` e equivalentes |
| Composição | `composicao_top_acoes.xlsx`, `composicao_top_dividendos.xlsx`, `composicao_top_small_caps.xlsx`, `composicao_esg.xlsx` |
| Return attribution | `decomposicao_top_acoes.xlsx`, `decomposicao_top_dividendos.xlsx`, `decomposicao_top_small_caps.xlsx`, `decomposicao_esg.xlsx` |
| Lâmina Comercial | `Lâmina Comercial - Top Ações.pptx`, `Lâmina Comercial - Top Dividendos.pptx`, `Lâmina Comercial - Top Small Caps.pptx` |
| Prestação de Contas | `Prestação de Contas - Top Ações - Agosto 2026.pptx`, `Prestação de Contas - Top Dividendos - Agosto 2026.pptx` e `Prestação de Contas - Top Small Caps - Agosto 2026.pptx` |

Cada arquivo de return attribution tem duas abas: uma para o mês atual aberto e
outra para o mês anterior fechado.

Os arquivos `componentes_*_comp_mes_passado_mtd_atual.xlsx` usam os ativos e
pesos da composição anterior ao último rebalanceamento, mas calculam
`Desempenho no mês` com os preços do mês atual até a data mais recente (MTD).

Confira a data de modificação e abra pelo menos um arquivo de cada grupo. Os
nomes são fixos e os arquivos anteriores são sobrescritos.

Não são mais gerados `lamina_dados_*.xlsx` nem
`composicao_compacta_*.xlsx`.

Na Prestação de Contas são atualizados automaticamente:

- mês no canto superior direito da primeira página;
- tabela inicial de desempenho;
- waterfall de decomposição de retornos;
- tabela de composição da carteira;
- gráfico de desempenho base 100.

O processo foi feito para rodar no primeiro dia útil. A tabela de desempenho e
o waterfall usam o mês recém-encerrado e a composição que efetivamente vigorou
durante esse mês. A tabela de composição mostra a nova carteira do mês que está
começando. Para Small Caps, o benchmark usado é o SMLL; nas demais, Ibovespa.

Antes de salvar o PPT, o programa confere se o total do waterfall coincide com
o retorno mensal da primeira página. Uma diferença relevante interrompe a
geração em vez de publicar números contraditórios.

Os textos editoriais das páginas 1 e 2 não são alterados pelo programa. Abra o
PPT gerado, revise esses textos manualmente e salve antes de distribuir.

## 9. Gerar o e-mail

Somente depois que o primeiro comando terminar, execute:

```powershell
.\.venv\Scripts\python.exe email_generator.py
```

O arquivo será salvo em:

```text
\\xpdocs\Research\Equities\Estrategia\Carteiras\Carteiras de Ações XP\email_carteiras.msg
```

Para salvar e abrir o rascunho:

```powershell
.\.venv\Scripts\python.exe email_generator.py --display
```

Revise os campos destacados em amarelo antes de enviar manualmente.

## 10. Executar os testes

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

O resultado esperado termina em `OK`. Os testes não acessam a rede corporativa
nem criam e-mails.

## 11. Erros comuns

### `FileNotFoundError` ou caminho não encontrado

- confirme a VPN;
- tente abrir `\\xpdocs` no Explorador de Arquivos;
- confira se algum arquivo corporativo foi renomeado.

### `PermissionError`

Feche o Excel que está usando o arquivo de saída e tente novamente. Também pode
faltar permissão de gravação na pasta.

### A Lâmina Comercial aparece como `[PULADO]`

O template configurado não foi encontrado. Copie o nome exato do arquivo pelo
Explorador e atualize `COMMERCIAL_TEMPLATE_FILES` em
`xp_carteiras\monthly_config.py`.

### A Prestação de Contas aparece como `[PULADO]`

Confirme que o template indicado em `ACCOUNTABILITY_TEMPLATE_FILES` existe na
pasta corporativa ou em `templates` dentro do projeto.

### Erro ao gerar a Prestação de Contas

- feche o PPT de saída antes de executar;
- confirme que o Microsoft PowerPoint abre normalmente;
- reinstale as dependências para garantir que `pywin32` está disponível;
- não interrompa o processo enquanto os gráficos estiverem sendo atualizados.

### `ModuleNotFoundError`

Repita a instalação das dependências:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Erro relacionado ao Outlook ou `win32com`

- abra o Outlook manualmente pelo menos uma vez;
- confirme que a conta está configurada;
- reinstale as dependências;
- execute no Windows.

### `Worksheet named ... not found`

A estrutura de uma planilha de entrada mudou. Confira se as abas do arquivo
`Performance carteiras.xlsm` ainda têm os mesmos nomes.

## 12. Atualizar o projeto

Se o projeto foi clonado com Git:

```powershell
git pull
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Se recebeu um novo ZIP, extraia-o em uma nova pasta e refaça as etapas de
preparação do ambiente.
