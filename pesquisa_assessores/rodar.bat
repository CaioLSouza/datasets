@echo off
chcp 65001 >nul

REM  pushd, e nao "cd /d": o cmd.exe nao aceita caminho UNC como
REM  diretorio atual. Como esta pasta fica em \\xpdocs\..., um "cd /d"
REM  cairia em C:\Windows sem avisar e todos os caminhos relativos
REM  (src\, config\, _dados\) quebrariam. O pushd mapeia uma letra de
REM  unidade temporaria para o compartilhamento e entra nela.
pushd "%~dp0"
if errorlevel 1 (
    echo  Nao consegui acessar a pasta do script:
    echo    %~dp0
    echo  Confira se a rede esta acessivel.
    pause
    exit /b 1
)

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
popd
pause
exit /b %CODIGO%
