"""Peças compartilhadas do pipeline da Pesquisa de Assessores.

Duas coisas aqui você pode querer editar:

  CAMINHOS  -- onde estão os arquivos na rede
  BLOCOS    -- as perguntas recorrentes da pesquisa

Não existe endereço de célula neste arquivo. A saída do pipeline são TABELAS;
onde cada gráfico mora é problema do Excel, não do Python.
"""
from __future__ import annotations

import csv
import os
import re
import unicodedata
from datetime import datetime
from pathlib import Path

# --------------------------------------------------------------------------
# CAMINHOS
# --------------------------------------------------------------------------
REDE = Path(r'\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores')
AQUI = Path(__file__).resolve().parent

# Para testar sem a rede: defina PA_REDE apontando para uma pasta local com a
# mesma estrutura. Em produção não precisa existir.
if os.environ.get('PA_REDE'):
    REDE = Path(os.environ['PA_REDE'])

CAMINHOS = {
    'input_forms':   REDE / 'input_forms',
    'pa_principal':  REDE / 'PA Principal.xlsx',   # só como fonte do histórico
    'bases':         REDE / 'bases',
    'base_geral':    REDE / 'bases' / 'PA Base.xlsx',
    'charts_data':   REDE / 'bases' / 'PA Charts Data.xlsx',
    'charts_csv':    REDE / 'bases' / 'charts',
    'pa_charts':     REDE / 'PA Charts.xlsx',
    'registro':      AQUI / 'registro.csv',
    'congelado':     AQUI / 'historico_congelado.csv',
    'ibovespa':      AQUI / 'ibovespa.csv',
    'logs':          AQUI / '_logs',
    'saida':         AQUI / '_saida',
}

# Ondas até esta (inclusive) usam o valor publicado, não o recalculado.
# Baixar este número é o que republicaria a série corrigida -- é uma decisão de
# negócio, não técnica.
ULTIMA_ONDA_PUBLICADA = 202607

# Quantas ondas as tabelas de série temporal carregam, por padrão. Um bloco
# pode pedir outra janela com `ondas_serie` -- proximos_meses e alocacao_rv
# usam 18, que é o que o report mostra.
ONDAS_NA_SERIE = 36

# Quantas perguntas do mês a planilha comporta. Cada uma ganha a sua própria
# tabela (q_mes_1, q_mes_2, ...), sempre emitida -- vazia quando não há
# pergunta para o slot. Slot fixo é o que permite montar o gráfico uma vez e
# não refazer: num mês com uma pergunta só, os outros ficam em branco.
#
# Medido no histórico: 11 ondas tiveram 1 pergunta extra, 10 tiveram 2,
# 6 tiveram 3 e 2 tiveram 5. Se passar de SLOTS_Q_MES, o pipeline avisa quais
# ficaram de fora -- não trunca calado.
SLOTS_Q_MES = 5

# Janela usada nas perguntas de enumeração fixa (ordenar='natural') para
# decidir se uma alternativa ainda existe. Sem resposta nenhuma nas últimas
# ONDAS_VIVAS ondas, ela sai dos gráficos. Larga de propósito: aqui o risco é
# derrubar uma faixa que só ficou vazia num mês.
ONDAS_VIVAS = 12

# --------------------------------------------------------------------------
# BLOCOS -- as perguntas recorrentes
# --------------------------------------------------------------------------
# tipo     : unica | multipla | escala
# match    : padrões para achar a coluna no export do Forms. Normalizado
#            (sem acento, minúsculo), casa por trecho. Acrescente sem apagar
#            os antigos -- é o que mantém as ondas velhas casando.
# ordenar  : 'valor'   -> maior percentual primeiro (ranking)
#            'natural' -> ordem do registro (faixas, escalas, regiões)
BLOCOS = [
    dict(id='regiao', tipo='unica', ordenar='natural',
         match=[r'regiao do brasil']),

    # ondas_serie: o report usa 18 meses nesta série, não os 36 do padrão
    dict(id='alocacao_rv', tipo='unica', ordenar='natural', ondas_serie=18,
         match=[r'alocacao em renda variavel esta']),

    dict(id='proximos_meses', tipo='unica', ordenar='natural', ondas_serie=18,
         match=[r'cenario dos proximos meses']),

    dict(id='classes_ativos', tipo='multipla', ordenar='valor',
         match=[r'classes de ativos']),

    dict(id='pct_internacional', tipo='unica', ordenar='natural',
         match=[r'ja investem em ativos internacionais']),

    dict(id='interesse_internacional', tipo='multipla', ordenar='valor',
         match=[r'investimentos internacionais, seus clientes se interessam']),

    dict(id='riscos_bolsa', tipo='multipla', ordenar='valor',
         match=[r'maiores risco', r'maior risco para a bolsa']),

    dict(id='setores', tipo='multipla', ordenar='valor',
         match=[r'setores da bolsa']),

    dict(id='sentimento', tipo='escala', ordenar='natural',
         match=[r'escala de 0 a 10']),

    # safra_rolante: as faixas mudam de ano em ano. Faixa de safra antiga entra
    # na base geral e não para a rodada.
    dict(id='ibovespa_alvo', tipo='unica', ordenar='natural', safra_rolante=True,
         match=[r'ibovespa atinja']),

    dict(id='apetite_risco', tipo='multipla', ordenar='valor',
         match=[r'apetite a risco', r'apetite por risco']),
]
BLOCO_POR_ID = {b['id']: b for b in BLOCOS}

# Rótulos que absorvem o texto livre da opção "Outra" do Forms.
ROTULOS_CATCHALL = {'outra', 'outras', 'outro', 'outros', 'other', 'others'}

# Acima desta fração de respondentes, o que cai na "Outra" deixa de ser texto
# livre e passa a ser sinal de que o conjunto de alternativas da onda não é o
# do registro. Texto livre é fio d'água; enchente é conjunto trocado.
LIMITE_CATCHALL = 0.20

# Lixo confirmado na Raw Data antiga: placeholders do Forms que ficaram na
# planilha. Descartados do numerador E do denominador, com aviso no log.
# Só entra aqui o que foi conferido a olho -- não é para virar tapete.
LIXO = {'opcao 1', 'opcao 2', 'opcao 3', 'opcao 4', 'opcao 5', 'opcao 6',
        'option 1', 'option 2', 'option 3', 'option 4', 'option 5', 'option 6'}

# Estrutura da aba Base da PA Principal -- usada SÓ pelo congelar.py, que lê
# dela o histórico publicado. Nada mais no pipeline conhece esses endereços.
BASE_ANTIGA = dict(linha_onda=1, linha_resp=2, linha_data=4,
                   col_en=1, col_pt=3, primeira_col=4)


# --------------------------------------------------------------------------
# normalização de texto
# --------------------------------------------------------------------------
def normalizar(s) -> str:
    """Minúsculo, sem acento, sem espaço duplo, sem ';' nem espaço nas pontas.

    É o que torna o casamento imune às reescritas de rótulo do Forms: o ';'
    colado, o \xa0 no lugar do espaço, o acento perdido.
    """
    if s is None:
        return ''
    s = str(s).replace('\xa0', ' ').replace('\u200b', '')
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = s.lower().strip().strip(';').strip()
    return re.sub(r'\s+', ' ', s)


def tokens(celula) -> list[str]:
    """Separa uma célula de múltipla escolha em rótulos normalizados.

    Compara token a token, então sobra ou falta de ';' no fim é indiferente --
    é isto que corrige o erro de contagem do processo antigo.
    """
    if celula is None:
        return []
    return [t for t in (normalizar(p) for p in str(celula).split(';')) if t]


def onda_de(data: datetime) -> int:
    return data.year * 100 + data.month


def data_da_onda(onda: int) -> datetime:
    return datetime(onda // 100, onda % 100, 1)


def onda_anterior(onda: int, todas: list[int]) -> int | None:
    """A onda imediatamente anterior que existe de fato (pode faltar mês)."""
    antes = [o for o in todas if o < onda]
    return max(antes) if antes else None


def slug(texto: str) -> str:
    s = re.sub(r'[^a-z0-9]+', '_', normalizar(texto)).strip('_')
    return s[:52] or 'x'


def chave_faixa(rotulo: str) -> float | None:
    """Extrai o número que ordena um rótulo de faixa. None se não for faixa.

    Serve as perguntas de enumeração fixa cujo rótulo é um intervalo -- as
    faixas de alocação, de percentual de clientes, de pontos do Ibovespa, e a
    escala 0-10. Ordenar por este número em vez da posição no registro é o que
    garante faixa em ordem crescente **mesmo quando o conjunto muda**: quando
    as faixas do Ibovespa rolarem de ano, ou quando entrar uma faixa no meio,
    não há `ordem` para acertar à mão.

    Foi assim que apareceu o problema: na Base antiga, "0% a 10%" e
    "10% a 25%" tinham sido acrescentadas ABAIXO de "25% a 50%", então a ordem
    do registro saía 25-50, 0-10, 10-25, 50-75, 75-100.

    "Abaixo de X" e "Acima de X" são deslocados para fora do intervalo, senão
    "Abaixo de 150 mil" empataria com "Entre 150 mil e 170 mil".
    """
    n = normalizar(rotulo)
    numeros = re.findall(r'\d+(?:[.,]\d+)?', n.replace('.', ''))
    if not numeros:
        return None
    valor = float(numeros[0].replace(',', '.'))
    if n.startswith(('abaixo', 'menos', 'ate', 'below', 'under')):
        return valor - 0.5
    if n.startswith(('acima', 'mais', 'above', 'over')):
        return valor + 0.5
    return valor


def ordem_natural(registros: list[dict]) -> dict[str, float]:
    """alternativa_id -> chave de ordenação, para blocos ordenar='natural'.

    Usa o número da faixa quando TODOS os rótulos são faixas; caso contrário
    cai na ordem do registro (região, "planejam aumentar/diminuir" etc., que
    não têm número e cuja ordem é editorial).
    """
    chaves = {r['alternativa_id']: chave_faixa(r['rotulo_pt']) for r in registros}
    if chaves and all(v is not None for v in chaves.values()):
        return chaves
    return {r['alternativa_id']: float(r['ordem']) for r in registros}



def ondas_na_janela(ondas: list[int], corrente: int, meses: int) -> list[int]:
    """As ondas dos últimos `meses` meses contados a partir de `corrente`.

    Por mês, não por contagem de ondas: o histórico tem dois meses sem
    pesquisa (dez/2020 e dez/2021), então "as últimas 18 ondas" e "os últimos
    18 meses" divergem. O report fala em meses, e assim a janela não estica
    para trás quando falta uma edição.
    """
    ano, mes = corrente // 100, corrente % 100
    total = ano * 12 + (mes - 1) - (meses - 1)
    primeira = (total // 12) * 100 + (total % 12) + 1
    return [o for o in ondas if primeira <= o <= corrente]

# --------------------------------------------------------------------------
# registro.csv
# --------------------------------------------------------------------------
CAMPOS_REGISTRO = ['pergunta_id', 'alternativa_id', 'serie_id', 'ordem',
                   'rotulo_pt', 'rotulo_en', 'aliases', 'valor_num', 'ativa',
                   'no_grafico']


def ler_registro(caminho: Path | None = None) -> list[dict]:
    caminho = caminho or CAMINHOS['registro']
    with open(caminho, encoding='utf-8-sig', newline='') as fh:
        linhas = [r for r in csv.DictReader(fh) if r.get('pergunta_id')]
    for r in linhas:
        r['ordem'] = int(r['ordem'] or 0)
        r['ativa'] = (r.get('ativa', '1').strip() or '1') != '0'
        # no_grafico: a alternativa existe e é apurada, mas fica fora das
        # tabelas d_/s_. É o caso do "Outra" em classes_ativos, riscos_bolsa e
        # interesse_internacional -- o deck não mostra. Em apetite_risco
        # mostra, então lá fica.
        r['no_grafico'] = (r.get('no_grafico', '1').strip() or '1') != '0'
    return linhas


def escrever_registro(linhas: list[dict], caminho: Path | None = None) -> None:
    caminho = caminho or CAMINHOS['registro']
    with open(caminho, 'w', encoding='utf-8-sig', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=CAMPOS_REGISTRO)
        w.writeheader()
        for r in linhas:
            w.writerow({k: ('1' if r.get(k) is True else
                            '0' if r.get(k) is False else r.get(k, ''))
                        for k in CAMPOS_REGISTRO})


def catchall_por_bloco(registro: list[dict]) -> dict[str, str]:
    """pergunta_id -> alternativa_id que recebe o texto livre da "Outra".

    Sem isso, cada resposta digitada à mão ("infarto do Lula") pararia a
    rodada. Com isso, ela cai na "Outra" -- que é onde o processo antigo
    também a colocava.
    """
    out: dict[str, str] = {}
    for r in registro:
        if normalizar(r['rotulo_pt']) in ROTULOS_CATCHALL:
            out.setdefault(r['pergunta_id'], r['alternativa_id'])
    return out


def indice_de_rotulos(registro: list[dict]) -> dict[tuple[str, str], dict]:
    """(pergunta_id, rótulo normalizado) -> linha do registro.

    É por aqui que uma alternativa renomeada continua caindo na mesma série.

    A chave inclui a pergunta porque o mesmo rótulo significa coisas
    diferentes em blocos diferentes: "Não estão interessados" existe em
    interesse_internacional, classes_ativos e setores.
    """
    idx: dict[tuple[str, str], dict] = {}
    for r in registro:
        chaves = [r['rotulo_pt'], r.get('rotulo_en', '')]
        chaves += (r.get('aliases') or '').split('|')
        for k in chaves:
            k = normalizar(k)
            if k:
                idx.setdefault((r['pergunta_id'], k), r)
    return idx


# --------------------------------------------------------------------------
# log
# --------------------------------------------------------------------------
class Log:
    def __init__(self, nome: str):
        CAMINHOS['logs'].mkdir(parents=True, exist_ok=True)
        self.caminho = CAMINHOS['logs'] / nome
        self.linhas: list[str] = []

    def __call__(self, msg: str = '') -> None:
        print(msg)
        self.linhas.append(msg)

    def gravar(self) -> None:
        self.caminho.write_text('\n'.join(self.linhas) + '\n', encoding='utf8')
        print(f'\n>> log: {self.caminho}')
