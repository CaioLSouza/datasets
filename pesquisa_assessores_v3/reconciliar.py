"""Confere a saída contra o que foi publicado.

    python reconciliar.py

Lê bases\\PA Base Historica.xlsx e responde duas perguntas:

  BLOCO 1 -- nenhum número publicado mudou?
             Compara a coluna `valor` com `pct_publicado` nas ondas congeladas.
             Tem que dar 100%. É o teste que autoriza usar o pipeline.

  BLOCO 2 -- onde o recálculo discorda do publicado?
             Compara `pct_calculado` com `pct_publicado`. Divergência aqui NÃO
             é erro do pipeline -- é o retrato dos problemas da planilha antiga
             (contagem truncada por ';' e denominador errado). Serve para você
             decidir, um dia, se vale republicar.
"""
from __future__ import annotations

import sys
from collections import defaultdict

import openpyxl

from comum import CAMINHOS, ULTIMA_ONDA_PUBLICADA

TOLERANCIA = 5e-5          # 0,005 pp


def carregar():
    caminho = CAMINHOS['base_geral']
    if not caminho.exists():
        print(f'ERRO: não achei {caminho}. Rode atualizar.py primeiro.')
        raise SystemExit(1)
    ws = openpyxl.load_workbook(caminho, data_only=True, read_only=True)['historico']
    it = ws.iter_rows(values_only=True)
    cab = list(next(it))
    return [dict(zip(cab, r)) for r in it if r[0] is not None]


def num(v):
    return None if v in (None, '') else float(v)


def main() -> int:
    linhas = carregar()
    print(f'histórico: {len(linhas)} valores\n')

    # ---------------------------------------------------------------- bloco 1
    print('=' * 70)
    print('BLOCO 1 -- o publicado saiu intacto?')
    print('=' * 70)
    conf = iguais = difs = 0
    exemplos = []
    for l in linhas:
        pub = num(l.get('pct_publicado'))
        if pub is None or int(l['onda']) > ULTIMA_ONDA_PUBLICADA:
            continue
        conf += 1
        if abs(float(l['valor']) - pub) <= TOLERANCIA:
            iguais += 1
        else:
            difs += 1
            if len(exemplos) < 10:
                exemplos.append(l)
    if conf:
        print(f'  valores congelados conferidos: {conf}')
        print(f'  idênticos ao publicado:        {iguais}  '
              f'({100 * iguais / conf:.2f}%)')
        print(f'  diferentes:                    {difs}')
        for l in exemplos:
            print(f'    onda {l["onda"]} {l["pergunta_id"]}/{l["alternativa_id"]}: '
                  f'saída={float(l["valor"]):.6f} publicado={num(l["pct_publicado"]):.6f}')
        print()
        print('  >> OK: nenhum número publicado mudou.' if difs == 0
              else '  >> ATENÇÃO: há número publicado divergente. Não use ainda.')
    else:
        print('  nada a conferir.')

    # ---------------------------------------------------------------- bloco 2
    print()
    print('=' * 70)
    print('BLOCO 2 -- onde o recálculo discorda do publicado?')
    print('=' * 70)
    pares = [l for l in linhas
             if num(l.get('pct_calculado')) is not None
             and num(l.get('pct_publicado')) is not None]
    if not pares:
        print('  nenhuma onda tem publicado e recalculado ao mesmo tempo.')
        print('  (rode com --bootstrap para trazer as ondas com dado bruto)')
        print()
        return 0

    por_pergunta: dict[str, list[float]] = defaultdict(list)
    grandes = []
    for l in pares:
        d = num(l['pct_calculado']) - num(l['pct_publicado'])
        por_pergunta[l['pergunta_id']].append(d)
        if abs(d) > 0.02:
            grandes.append((abs(d), d, l))

    print(f'  pares publicado x recalculado: {len(pares)}')
    print(f'  batem (até 0,005 pp):          '
          f'{sum(1 for l in pares if abs(num(l["pct_calculado"]) - num(l["pct_publicado"])) <= TOLERANCIA)}')
    print(f'  divergem mais de 2 pp:         {len(grandes)}')
    print()
    print('  por pergunta (erro médio absoluto, em pp):')
    for pid, ds in sorted(por_pergunta.items(),
                          key=lambda kv: -sum(abs(d) for d in kv[1]) / len(kv[1])):
        mae = 100 * sum(abs(d) for d in ds) / len(ds)
        pior = 100 * max(ds, key=abs)
        print(f'    {pid:<26} n={len(ds):>4}  mae={mae:>6.2f}  pior={pior:>+7.2f}')

    grandes.sort(key=lambda t: t[0], reverse=True)
    if grandes:
        print()
        print('  as 12 maiores divergências:')
        for _, d, l in grandes[:12]:
            print(f'    onda {l["onda"]} {l["pergunta_id"]}/{l["alternativa_id"][:34]:<34} '
                  f'pub={100 * num(l["pct_publicado"]):>6.2f}%  '
                  f'calc={100 * num(l["pct_calculado"]):>6.2f}%  '
                  f'{100 * d:>+7.2f} pp')
    print()
    print('  Nada disso alimenta gráfico: as ondas congeladas usam o publicado.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
