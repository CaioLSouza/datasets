# Tutorial para iniciantes

Este tutorial ensina a executar a automação no notebook da empresa. Não é
necessário saber programar.

## 1. O que o projeto faz

O processo possui dois comandos:

1. `portfolio_automation.py` lê as bases corporativas, calcula as carteiras,
   gera os Excel e atualiza as Lâminas Comerciais.
2. `email_generator.py` usa esses Excel para criar `email_carteiras.msg`.

O projeto mantém a **Lâmina Comercial**, mas não gera a lâmina completa
auxiliar (`lamina_dados_*.xlsx`) nem a Prestação de Contas. O e-mail não é
enviado automaticamente: ele é salvo para revisão.

## 2. Antes de começar

Confirme:

- Windows conectado à rede da empresa ou VPN;
- acesso às pastas `\\xpdocs`;
- Microsoft Outlook instalado e configurado;
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

## 6. Conferir o template da Lâmina Comercial

Abra `xp_carteiras\monthly_config.py` e confira os nomes em
`COMMERCIAL_TEMPLATE_FILES`. Eles devem ser idênticos aos arquivos existentes
na pasta de templates, incluindo mês, ano, acentos, espaços e `.pptx`.

Altere somente o nome à direita dos dois-pontos. Não altere o nome da carteira.

## 7. Executar os cálculos, Excel e Lâmina Comercial

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
| Componentes | `componentes_top_acoes_atual.xlsx`, `componentes_top_acoes_ultimo_rebal.xlsx` e equivalentes |
| Composição | `composicao_top_acoes.xlsx`, `composicao_top_dividendos.xlsx`, `composicao_top_small_caps.xlsx`, `composicao_esg.xlsx` |
| Return attribution | `decomposicao_top_acoes.xlsx`, `decomposicao_top_dividendos.xlsx`, `decomposicao_top_small_caps.xlsx`, `decomposicao_esg.xlsx` |
| Lâmina Comercial | `Lâmina Comercial - Top Ações.pptx`, `Lâmina Comercial - Top Dividendos.pptx`, `Lâmina Comercial - Top Small Caps.pptx` |

Cada arquivo de return attribution tem duas abas: uma para o mês atual aberto e
outra para o mês anterior fechado.

Confira a data de modificação e abra pelo menos um arquivo de cada grupo. Os
nomes são fixos e os arquivos anteriores são sobrescritos.

Não são mais gerados `lamina_dados_*.xlsx`,
`composicao_compacta_*.xlsx` nem PowerPoint de Prestação de Contas.

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
