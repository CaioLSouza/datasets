"""Gera o montar_charts.py embutindo o montar_charts.ps1.

Roda quando o .ps1 mudar:

    python _gerar_montar_charts.py

Por que existe: o Gmail se recusa a anexar .ps1 (política de executáveis) e o
GitHub está bloqueado na máquina corporativa. Um .py único passa pelos dois
caminhos, então o PowerShell viaja embutido dentro dele.

O embutimento é feito por código, nunca à mão -- é o que garante que o
PowerShell chegue byte a byte igual ao .ps1 testado.
"""
from pathlib import Path

AQUI = Path(__file__).resolve().parent
PS1 = AQUI / 'montar_charts.ps1'
SAIDA = AQUI / 'montar_charts.py'

Q3 = '"' * 3          # os delimitadores, montados para não colidir com o texto
ABRE = 'r' + Q3

ps = PS1.read_text(encoding='utf8')
if Q3 in ps:
    raise SystemExit('ERRO: o .ps1 contém """ -- o embutimento precisaria escapar')
if ps.rstrip().endswith('\\'):
    raise SystemExit('ERRO: o .ps1 termina em barra invertida -- quebraria o r-string')

CABECALHO = '\n'.join([
    Q3 + 'Cria a PA Charts.xlsx -- a planilha onde você monta os gráficos.',
    '',
    '    python montar_charts.py',
    '    python montar_charts.py --dados "...\\\\bases\\\\charts" --saida "...\\\\PA Charts.xlsx"',
    '',
    'Roda UMA vez, na instalação. Depois disso a planilha é sua: você insere os',
    'gráficos, formata, e todo mês é só abrir e apertar Atualizar Tudo.',
    '',
    'O trabalho quem faz é PowerShell, porque criar consulta Power Query exige',
    'COM do Excel e o pywin32 não está instalado aqui. O script PowerShell vive',
    'embutido neste arquivo: ele o grava num temporário, executa e apaga.',
    '',
    'Não edite o PowerShell aqui dentro. Edite o montar_charts.ps1 e rode',
    '_gerar_montar_charts.py -- assim o que roda é sempre o que foi testado.',
    Q3,
    'from __future__ import annotations',
    '',
    'import subprocess',
    'import sys',
    'import tempfile',
    'from pathlib import Path',
    '',
    'sys.path.insert(0, str(Path(__file__).resolve().parent))',
    'from comum import CAMINHOS',
    '',
    '# --------------------------------------------------------------------------',
    '# O montar_charts.ps1, na íntegra. Gerado -- não edite.',
    '# --------------------------------------------------------------------------',
    'POWERSHELL = ' + ABRE,
])

RODAPE = '\n'.join([
    Q3,
    '',
    '',
    'def main() -> int:',
    '    args = sys.argv[1:]',
    "    dados = CAMINHOS['charts_csv']",
    "    saida = CAMINHOS['pa_charts']",
    '    for i, a in enumerate(args):',
    "        if a == '--dados' and i + 1 < len(args):",
    '            dados = Path(args[i + 1])',
    "        elif a == '--saida' and i + 1 < len(args):",
    '            saida = Path(args[i + 1])',
    '',
    '    if not dados.exists():',
    "        print(f'ERRO: não achei a pasta de dados:')",
    "        print(f'  {dados}')",
    "        print('Rode o atualizar.py primeiro.')",
    '        return 1',
    '    if saida.exists():',
    "        print(f'ERRO: {saida} já existe.')",
    "        print('Este script cria a planilha do zero e apagaria os seus')",
    "        print('gráficos. Se é isso que você quer, renomeie ou mova a atual.')",
    '        return 1',
    '',
    '    # o temporário fica local, não na rede -- o PowerShell carrega mais rápido',
    '    with tempfile.TemporaryDirectory() as tmp:',
    "        ps1 = Path(tmp) / 'montar_charts.ps1'",
    "        ps1.write_text(POWERSHELL, encoding='utf-8-sig')",
    "        print(f'PowerShell embutido: {len(POWERSHELL)} caracteres')",
    "        print('isto leva 3 a 4 minutos -- o Excel monta consulta por consulta')",
    "        print()",
    '        proc = subprocess.run(',
    "            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',",
    "             '-File', str(ps1), '-Dados', str(dados), '-Saida', str(saida)],",
    '            text=True)',
    '    return proc.returncode',
    '',
    '',
    "if __name__ == '__main__':",
    '    raise SystemExit(main())',
    '',
])

# Sem newline entre o r""" e o PowerShell: o primeiro caractere depois do
# delimitador já faz parte da string, então um '\n' aqui viraria uma linha em
# branco no começo do script gravado.
SAIDA.write_text(CABECALHO + ps + RODAPE, encoding='utf8')
print(f'>> {SAIDA.name}  ({SAIDA.stat().st_size} bytes)')

# --- confere que o PowerShell embutido é idêntico ao .ps1
import ast
import re
src = SAIDA.read_text(encoding='utf8')
ast.parse(src)
print('   sintaxe Python: ok')
mod = ast.parse(src)
embutido = next(n.value.value for n in ast.walk(mod)
                if isinstance(n, ast.Assign)
                and getattr(n.targets[0], 'id', '') == 'POWERSHELL')
print(f'   PowerShell embutido idêntico ao .ps1: {embutido == ps}')
if embutido != ps:
    raise SystemExit('ERRO: o embutimento não preservou o PowerShell')
