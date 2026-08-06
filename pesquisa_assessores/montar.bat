@echo off
chcp 65001 >nul

REM  enabledelayedexpansion e obrigatorio aqui: sem ele, o %RESP% dentro
REM  do bloco if seria expandido na hora em que o cmd LE o bloco, antes
REM  do "set /p" rodar -- e a confirmacao nunca funcionaria. Com ele,
REM  usa-se !RESP!, que e lido na hora da execucao.
setlocal enabledelayedexpansion

REM =====================================================================
REM  MONTAGEM — roda UMA VEZ, na virada para o processo novo
REM =====================================================================
REM  Faz as tres fases em sequencia, parando na primeira que falhar:
REM
REM     1. congelar_historico   le a aba Base e grava o que foi publicado
REM     2. pipeline --bootstrap le a aba Raw Data e importa as respostas
REM     3. reconciliar          confere que nada mudou de valor
REM
REM  Para a rodada de todo mes, use o rodar.bat -- nao este.
REM
REM  pushd em vez de "cd /d": o cmd.exe nao aceita caminho UNC como
REM  diretorio atual, e esta pasta pode ficar em \\xpdocs\...
REM =====================================================================

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
echo   PESQUISA DE ASSESSORES XP - MONTAGEM
echo   (roda uma vez so)
echo  ===============================================
echo.

REM ---------------------------------------------------------------------
REM  Trava: recongelar por engano reescreveria o historico publicado.
REM  Enquanto a PA Principal estiver intocada o resultado e o mesmo, mas
REM  se alguem tiver mexido nela, os numeros "congelados" mudariam em
REM  silencio -- exatamente o que este desenho existe para impedir.
REM ---------------------------------------------------------------------
if exist "config\valores_publicados.csv" (
    echo  ATENCAO: config\valores_publicados.csv ja existe.
    echo.
    echo  O historico ja foi congelado antes. Rodar de novo vai
    echo  reescrever esse arquivo a partir da PA Principal de hoje.
    echo.
    echo  Se a PA Principal nao mudou, o resultado sera identico.
    echo  Se alguem mexeu nela, os numeros publicados vao mudar.
    echo.
    set "RESP="
    set /p RESP="  Continuar mesmo assim? (digite SIM) "
    if /i not "!RESP!"=="SIM" (
        echo.
        echo  Cancelado. Nada foi alterado.
        popd
        pause
        exit /b 0
    )
    echo.
)

echo  -----------------------------------------------
echo   FASE 1 de 3 - congelando o historico publicado
echo  -----------------------------------------------
python src\congelar_historico.py %*
if errorlevel 1 goto :falhou

echo.
echo  -----------------------------------------------
echo   FASE 2 de 3 - importando as respostas
echo  -----------------------------------------------
python src\pipeline.py --bootstrap %*
if errorlevel 1 goto :falhou

echo.
echo  -----------------------------------------------
echo   FASE 3 de 3 - conferindo que nada mudou
echo  -----------------------------------------------
python src\reconciliar.py %*
if errorlevel 1 goto :falhou

echo.
echo  ===============================================
echo   MONTAGEM CONCLUIDA
echo  ===============================================
echo.
echo   Confira acima, no bloco 1 da conferencia:
echo     - IDENTICOS tem que dar ~100%% (hoje: 3717 de 3718)
echo     - "Ondas na Base" tem que dizer 76
echo.
echo   Se bateu, o proximo passo e montar a PA Report.xlsx.
echo   O passo a passo esta em powerquery\INSTRUCOES.md
echo.
popd
pause
exit /b 0

:falhou
echo.
echo  ***********************************************
echo   A MONTAGEM PAROU. Nada foi gravado.
echo   Leia a mensagem acima: ela diz o que houve.
echo  ***********************************************
echo.
popd
pause
exit /b 1
