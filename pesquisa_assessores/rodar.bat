@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo  ===============================================
echo   PESQUISA DE ASSESSORES XP - rodada mensal
echo  ===============================================
echo.

python src\pipeline.py %*
set CODIGO=%ERRORLEVEL%

echo.
if %CODIGO% NEQ 0 (
    echo  ***********************************************
    echo   A RODADA NAO FOI CONCLUIDA. Nada foi gravado.
    echo   Leia o log acima: ele diz o que precisa ser
    echo   ajustado em config\perguntas.yaml.
    echo  ***********************************************
) else (
    echo  Pronto. Agora:
    echo    1^) abra a PA Report.xlsx
    echo    2^) Dados ^> Atualizar Tudo
    echo    3^) abra os dois PPTs e Atualizar Links
)
echo.
pause
exit /b %CODIGO%
