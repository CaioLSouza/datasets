# Instalação sem internet no pip

Estes três arquivos existem por um motivo específico: na rede da XP o
`pip` não alcança o `pypi.org`, e tentar instalar direto devolve
`No matching distribution found` mesmo com o nome do pacote correto.

O GitHub costuma passar onde o PyPI é bloqueado. Por isso as
dependências estão aqui.

---

## Instalar

```bash
py -m pip install --no-index --find-links _wheels pyyaml openpyxl
```

Rode de dentro da pasta `pesquisa_assessores`. Se preferir caminho
absoluto, troque `_wheels` pelo caminho completo da pasta.

- `--no-index` impede que o pip tente o pypi.org e falhe de novo
- `--find-links` manda procurar nesta pasta
- o `et-xmlfile` é resolvido sozinho: ele é dependência do openpyxl e
  está aqui junto

Confira depois:

```bash
py -c "import yaml, openpyxl; print('ok')"
```

**Use `py`, não `python`.** Na máquina corporativa o `py` aponta para o
Python gerenciado em `C:\Users\Public\Programs\`. Se `python` resolver
para outro interpretador, você instala num e roda no outro — e o erro
continua igual.

---

## O que são

Baixados do PyPI com `pip download`, sem modificação. Todos com licença
MIT, redistribuição permitida.

| Pacote | Versão | Para quê |
|---|---|---|
| PyYAML | 6.0.3 | ler `config/config.yaml` e `config/perguntas.yaml` |
| openpyxl | 3.1.5 | ler e escrever `.xlsx` |
| et-xmlfile | 2.0.0 | dependência do openpyxl |

---

## Compatibilidade

O wheel do PyYAML é **compilado**, então serve para um alvo só:

```
pyyaml-6.0.3-cp310-cp310-win_amd64.whl
              └─┬─┘        └────┬───┘
           Python 3.10     Windows 64 bits
```

É o que roda na máquina corporativa (Python 3.10.4, 64 bits). **Em outra
versão de Python isso não instala** — o pip vai dizer que o wheel não é
suportado nesta plataforma.

Se um dia o Python mudar de versão, gere os novos assim, numa máquina
com internet:

```bash
python -m pip download pyyaml openpyxl --dest _wheels --python-version 3.12 --platform win_amd64 --only-binary=:all:
```

Trocando `3.12` pela versão de destino. Os outros dois (`openpyxl` e
`et-xmlfile`) são `py3-none-any` — funcionam em qualquer versão, não
precisam ser regerados.

---

## Conferir a integridade

Antes de instalar binário vindo da internet numa máquina corporativa,
vale checar. No PowerShell:

```powershell
Get-FileHash _wheels\*.whl -Algorithm SHA256 | Format-List Path, Hash
```

Tem que bater com:

```
et_xmlfile-2.0.0-py3-none-any.whl
  7A91720BC756843502C3B7504C77B8FE44217C85C537D85037F0F536151B2CAA

openpyxl-3.1.5-py2.py3-none-any.whl
  5282C12B107BFFEEF825F4617DC029AFAF41D0EA60823BBB665EF3079DC79DE2

pyyaml-6.0.3-cp310-cp310-win_amd64.whl
  BDB2C67C6C1390B63C6FF89F210C8FD09D9A1217A465701EAC7316313C915E4C
```

Os mesmos valores estão publicados nas páginas dos pacotes no PyPI, em
*Download files → view hashes*, se quiser confrontar com a fonte.

---

## O consertozinho definitivo

Isto aqui é contorno. Empresa que configura certificado próprio — e a
XP configura, veja o `xpi-ca-bundle.pem` no `pip config list` — quase
sempre mantém um espelho interno do PyPI (Artifactory ou Nexus).

Vale abrir um chamado perguntando: **"qual o index-url do mirror interno
do PyPI?"** Com ele, o `pip install` volta a funcionar normalmente e
ninguém precisa carregar wheel na mão na próxima vez.
