@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo  Pesquisa de Assessores -- atualizacao mensal
echo ============================================================
echo.

python atualizar.py
if errorlevel 1 goto :parou

echo.
python reconciliar.py
if errorlevel 1 goto :parou

echo.
echo ============================================================
echo  PRONTO.
echo.
echo  Falta so abrir a PA Charts.xlsx e usar
echo    Dados ^> Atualizar Tudo
echo.
echo  Os graficos das perguntas recorrentes se atualizam sozinhos.
echo  O da pergunta do mes e o unico que voce remonta -- os numeros
echo  dele estao na aba q_mes.
echo ============================================================
echo.
pause
exit /b 0

:parou
echo.
echo ============================================================
echo  PAROU. Nada foi gravado.
echo.
echo  A mensagem acima diz o que precisa de atencao. O log fica
echo  em _logs\ . O caso mais comum e alternativa nova na
echo  pesquisa: e uma linha no registro.csv.
echo ============================================================
echo.
pause
exit /b 1
