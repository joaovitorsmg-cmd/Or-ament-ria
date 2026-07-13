@echo off
chcp 65001 > nul
title Agroquima -- Gerador de Painel Orcamentario

echo.
echo  =============================================
echo   Agroquima -- Gerador de Painel Orcamentario
echo  =============================================
echo.
echo  Arquivos necessarios nesta pasta:
echo    - Mapa_Despesa.xlsx
echo    - Base_regionais.xlsx
echo    - gerar_json.py
echo.

:: Verifica se Python esta instalado
python --version > nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  [ERRO] Python nao encontrado no PATH do sistema.
    echo.
    echo  Como instalar:
    echo    1. Acesse https://www.python.org/downloads/
    echo    2. Baixe a versao mais recente (3.10 ou superior)
    echo    3. Na instalacao, MARQUE a opcao "Add Python to PATH"
    echo    4. Feche e reabra esta janela apos instalar
    echo.
    pause
    exit /b 1
)

:: Verifica se pandas esta instalado
python -c "import pandas" > nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  [INFO] Instalando dependencias necessarias (pandas, openpyxl)...
    echo         Aguarde, isso so ocorre na primeira vez.
    echo.
    python -m pip install pandas openpyxl --quiet
    if %ERRORLEVEL% NEQ 0 (
        echo  [ERRO] Falha ao instalar dependencias.
        echo         Tente manualmente: pip install pandas openpyxl
        pause
        exit /b 1
    )
    echo  [OK] Dependencias instaladas com sucesso.
    echo.
)

echo  Gerando aqm_data.json...
echo.

python gerar_json.py

echo.
if %ERRORLEVEL% NEQ 0 (
    echo  -----------------------------------------------
    echo  [ERRO] A geracao falhou. Leia as mensagens acima
    echo         para entender o problema.
    echo  -----------------------------------------------
) else (
    echo  -----------------------------------------------
    echo  [OK] Arquivo aqm_data.json gerado com sucesso!
    echo       Arraste-o no painel (aba Atualizar Base).
    echo  -----------------------------------------------
)
echo.
pause
