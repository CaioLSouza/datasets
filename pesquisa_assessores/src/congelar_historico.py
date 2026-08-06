# -*- coding: utf-8 -*-
"""
=======================================================================
 CONGELADOR DO HISTORICO  (roda UMA vez, na migracao)
=======================================================================

 Numero que ja foi publicado nao muda. Este script le a aba Base da PA
 Principal e grava os valores COMO ESTAO PUBLICADOS hoje, onda a onda.

 A partir dai o pipeline passa a funcionar assim:

   ondas ate `ultima_onda_publicada`  -> valor CONGELADO (o publicado)
   ondas novas                        -> valor CALCULADO do bruto

 Os merges de alias continuam valendo, e nao mexem em numero nenhum:
 como "Riscos geopoliticos" (ate 202603) e "Riscos geopoliticos/ Guerra"
 (a partir de 202604) vivem em ondas disjuntas, juntar as duas na mesma
 linha so emenda a serie - os valores de cada mes seguem identicos ao
 que foi publicado.

 SAIDA
   config/valores_publicados.csv   onda;chave;pct
       O congelado. E este arquivo que o pipeline le para nao
       recalcular o que ja foi publicado. Versione junto com o
       registro de perguntas.

   _saida/chaves_base.csv          linha;chave;texto;reconhecido_como
       O mapa de como cada linha da Base antiga foi interpretada.
       Nao e usado pelo pipeline -- serve para voce auditar o
       congelamento e para achar os blocos mortos da planilha velha.

 USO
   python congelar_historico.py "\\\\xpdocs\\...\\PA Principal.xlsx"
=======================================================================
"""

import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import openpyxl

from pipeline import RAIZ, Registro, norm

IGNORAR = {"total", "media", "média", "resposta media", "resposta média", ""}


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    pa = Path(sys.argv[1])
    reg = Registro(RAIZ / "config" / "perguntas.yaml")

    wb = openpyxl.load_workbook(pa, data_only=True)
    ws = wb["Base"]

    # colunas de onda: codigo AAAAMM na linha 1, nº de respostas na linha 2
    ondas, bases = {}, {}
    for c in range(4, 200):
        v = ws.cell(1, c).value
        if isinstance(v, (int, float)) and 202000 < v < 210000:
            ondas[c] = int(v)
            b = ws.cell(2, c).value
            bases[c] = float(b) if isinstance(b, (int, float)) and b else None

    linhas_mapa, congelado = [], []
    atual, atual_safra, buffer = None, None, []
    stats = {"bloco": 0, "opcao": 0, "orfa": 0, "contagem": 0}

    def fechar_bloco():
        """Grava o bloco acumulado, decidindo % x contagem por COLUNA.

        A Base guarda a escala 0-10 em contagem num bloco e em % noutro.
        A decisao nao pode ser por celula (contagem de 1 pessoa seria
        confundida com 100%) nem pela soma (em multipla escolha os
        percentuais somam bem mais que 1).

        A regra que vale para os dois casos: percentual nunca passa de 1.
        Se o MAIOR valor da coluna dentro do bloco estoura 1, o bloco
        esta em contagem.
        """
        if not buffer:
            return
        for col, onda in ondas.items():
            vals = [(ch, ws.cell(r, col).value) for r, ch in buffer]
            nums = [v for _, v in vals if isinstance(v, (int, float))]
            if not nums:
                continue
            divisor = 1.0
            if max(nums) > 1.05:                # bloco em contagem
                divisor = bases.get(col) or 0
                if not divisor:
                    continue
                stats["contagem"] += 1
            era_pct = divisor == 1.0
            for ch, v in vals:
                if isinstance(v, (int, float)):
                    congelado.append((onda, ch, v / divisor, era_pct))

    for r in range(1, ws.max_row + 1):
        c = ws.cell(r, 3).value
        txt = str(c).strip() if c not in (None, "") else ""
        if norm(txt) in IGNORAR:
            linhas_mapa.append((r, "", txt[:60], "—"))
            continue
        p, safra = reg.identificar(txt)
        if p:
            fechar_bloco()
            buffer = []
            atual, atual_safra = p, safra
            stats["bloco"] += 1
            marca = f" safra {safra}" if safra else ""
            linhas_mapa.append((r, "", txt[:60], f"cabeçalho de {p.id}{marca}"))
            continue
        if atual is None:
            linhas_mapa.append((r, "", txt[:60], "fora de bloco conhecido"))
            continue
        res = atual.resolver(txt)
        if not res or res[0][3] == "desconhecida":
            fechar_bloco()
            buffer = []
            atual = None
            stats["orfa"] += 1
            linhas_mapa.append((r, "", txt[:60], "não reconhecida — bloco encerrado"))
            continue

        chave = (f"{atual.id}|{atual_safra}|{res[0][0]}" if atual_safra
                 else f"{atual.id}|{res[0][0]}")
        stats["opcao"] += 1
        linhas_mapa.append((r, chave, txt[:60], atual.id))

        # Chave repetida dentro do mesmo bloco = comecou um sub-bloco.
        # A Base empilha, sem cabecalho no meio, a versao em CONTAGEM e a
        # versao em % da escala 0-10 (o "Resposta media" entre as duas
        # nao fecha nada). Sem este corte, os dois viram um bloco so e a
        # deteccao de escala se perde.
        if any(ch == chave for _, ch in buffer):
            fechar_bloco()
            buffer = []

        buffer.append((r, chave))

    fechar_bloco()

    # --------- grava ---------
    (RAIZ / "config").mkdir(exist_ok=True)
    (RAIZ / "_saida").mkdir(exist_ok=True)

    # Uma mesma chave pode vir de DUAS linhas da Base: o rotulo antigo e
    # o novo da alternativa renomeada (ex.: "Instabilidade politica" e
    # "Instabilidade politica/ Eleicoes"). Como as janelas sao disjuntas,
    # em cada onda so uma das duas linhas tem valor de verdade e a outra
    # esta zerada. Fico com a que tem valor.
    #  Regra de desempate, nesta ordem:
    #   1. valor nao-zerado ganha do zerado  (rotulo renomeado: em cada
    #      onda so uma das duas linhas tem valor)
    #   2. bloco que ja estava em % ganha do bloco em contagem  (a escala
    #      0-10 aparece nos dois formatos; o publicado como % e o que
    #      saiu no report)
    #  Sobra conflito de verdade so quando duas linhas em % discordam —
    #  ai provavelmente ha alias errado no perguntas.yaml.
    dest_val = RAIZ / "config" / "valores_publicados.csv"
    melhor, conflitos = {}, []
    for onda, chave, v, era_pct in congelado:
        k = (onda, chave)
        if k not in melhor:
            melhor[k] = (v, era_pct)
            continue
        atual_v, atual_pct = melhor[k]
        if atual_v in (0, None) and v not in (0, None):
            melhor[k] = (v, era_pct)
        elif v in (0, None):
            pass
        elif era_pct and not atual_pct:
            melhor[k] = (v, era_pct)
        elif era_pct == atual_pct and abs(atual_v - v) > 0.0015:
            # Duas linhas em % discordando. Na Base, correcao entra como
            # linha NOVA logo abaixo da errada (foi o que aconteceu com
            # "Melhora na recuperacao economica global": a de cima tem um
            # ';' colado no rotulo e subconta). Fico com a de baixo e
            # aviso, para voce conferir e, se quiser, editar o CSV na mao.
            conflitos.append((onda, chave, atual_v, v))
            melhor[k] = (v, era_pct)

    with open(dest_val, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["onda", "chave", "pct"])
        for (onda, chave), (v, _) in sorted(melhor.items()):
            w.writerow([onda, chave, repr(v)])
    vistos = melhor

    dest_map = RAIZ / "_saida" / "chaves_base.csv"
    with open(dest_map, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write("linha;chave;texto_na_coluna_C;reconhecido_como\n")
        for r, ch, txt, como in linhas_mapa:
            fh.write(f"{r};{ch};{txt.replace(';', ',')};{como}\n")

    ondas_u = sorted({o for o, _ in melhor})
    print(f"Blocos reconhecidos ....... {stats['bloco']}")
    print(f"Linhas de alternativa ..... {stats['opcao']}")
    print(f"Linhas fora do registro ... {stats['orfa']}  (perguntas do mês antigas)")
    print(f"\nValores congelados ........ {len(vistos)}")
    print(f"Ondas cobertas ............ {len(ondas_u)}  ({ondas_u[0]}..{ondas_u[-1]})")
    if conflitos:
        print(f"\nATENÇÃO: {len(conflitos)} chave(s) com duas linhas na Base "
              f"publicando valores diferentes.")
        print("Fiquei com a linha de baixo (na Base, correção entra abaixo da")
        print("linha errada). Confira; se preferir a outra, edite o valor à mão")
        print("em config/valores_publicados.csv — é um CSV simples.")
        print(f"\n   {'onda':>7}  {'chave':<40} {'de cima':>9} {'DE BAIXO':>9}")
        for onda, chave, a, b in conflitos[:10]:
            print(f"   {onda:>7}  {chave:<40} {a:>9.4f} {b:>9.4f}")
    print(f"\nGerado: {dest_val}")
    print(f"Gerado: {dest_map}")
    print(f"\nAgora ponha em config/config.yaml:")
    print(f"    ultima_onda_publicada: {ondas_u[-1]}")
    print("A partir da onda seguinte, o pipeline calcula do bruto.")


if __name__ == "__main__":
    main()
