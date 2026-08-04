# Automação das carteiras XP

Este projeto calcula as carteiras, gera os arquivos Excel operacionais,
atualiza as **Lâminas Comerciais**, gera a **Prestação de Contas da Top Ações**
e cria o e-mail mensal no Outlook. Os Excel auxiliares
`lamina_dados_*.xlsx` continuam removidos.

Para instruções detalhadas, consulte o
[tutorial para iniciantes](TUTORIAL_INICIANTE.md).

## Comandos principais

- `portfolio_automation.py`: calcula performance, componentes, composições e
  return attribution, exporta os Excel e atualiza os PowerPoint.
- `email_generator.py`: usa os Excel gerados para montar e salvar
  `email_carteiras.msg`.

## Organização interna

```text
portfolio_automation.py          comando do pipeline principal
email_generator.py               comando de geração do e-mail
xp_carteiras/
  pipeline.py                    ordem das etapas e exportações
  performance.py                 retornos, risco e estatísticas
  components.py                  composição e atribuição por papel
  excel_reports.py               criação e formatação dos Excel
  powerpoint_reports.py          atualização da Lâmina Comercial
  accountability_reports.py      geração da Prestação de Contas
  email_report.py                cálculos, HTML e Outlook
  monthly_config.py              nomes mensais dos templates comerciais
  settings.py                    caminhos de entrada e saída
  constants.py                   nomes e mapas estáticos
templates/                        template empacotado da Prestação de Contas
tests/                            testes com dados sintéticos
```

Importar os módulos não executa o pipeline. Leituras da rede corporativa,
gravações e acesso ao Outlook acontecem somente quando os comandos principais
são executados.

## Instalação

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Os caminhos padrão continuam apontando para `\\xpdocs`.

A Prestação de Contas requer Windows com Microsoft PowerPoint instalado. O
template fornecido fica também em `templates`, como alternativa ao arquivo da
pasta corporativa.

## Execução

Execute primeiro:

```powershell
.\.venv\Scripts\python.exe portfolio_automation.py
```

Depois gere o e-mail:

```powershell
.\.venv\Scripts\python.exe email_generator.py
```

Para também abrir o rascunho no Outlook:

```powershell
.\.venv\Scripts\python.exe email_generator.py --display
```

O programa não envia o e-mail automaticamente.

## Configuração opcional

Os caminhos podem ser substituídos por variáveis de ambiente:

| Variável | Uso |
| --- | --- |
| `XP_PORTFOLIO_ROOT` | pasta raiz de Carteiras de Ações XP |
| `XP_CROSS_DATA_DIR` | pasta `_Cross Data` |
| `XP_OUTPUT_DIR` | saída dos Excel intermediários |
| `XP_EMAIL_DIR` | pasta de `email_carteiras.msg` |
| `XP_PERFORMANCE_WORKBOOK` | arquivo `Performance carteiras.xlsm` |
| `XP_COMP_SHEET_PATH` | arquivo `COMP SHEET/raw_data.xlsx` |
| `XP_TEMPLATES_DIR` | templates dos PowerPoint |
| `XP_COMMERCIAL_DECK_DIR` | saída das Lâminas Comerciais |
| `XP_ACCOUNTABILITY_DECK_DIR` | saída da Prestação de Contas |
| `XP_SECTOR_CLASSIFICATION_PATH` | classificação setorial |
| `XP_MARKET_DATA_PATH` | parquet de market data |
| `XP_BDR_MARKET_DATA_PATH` | CSV de BDRs |
| `XP_INDICES_PATH` | parquet de índices |
| `XP_INDEX_COMPOSITION_PATH` | composição do Ibovespa |

Exemplo:

```powershell
$env:XP_OUTPUT_DIR = "C:\temp\carteiras\output"
.\.venv\Scripts\python.exe portfolio_automation.py
```

## Testes

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Os testes usam dados sintéticos e não acessam `\\xpdocs` nem o Outlook.

## Escopo dos PowerPoint

O pipeline não gera mais:

- `lamina_dados_*.xlsx`;
- `composicao_compacta_*.xlsx`.

Os PowerPoint `Lâmina Comercial - *.pptx` continuam sendo atualizados. A
Prestação de Contas é gerada inicialmente para a Top Ações, com mês, tabela de
desempenho, waterfall, composição e gráfico base 100 atualizados. Os textos
editoriais das páginas 1 e 2 permanecem como estão no template para revisão
manual.
