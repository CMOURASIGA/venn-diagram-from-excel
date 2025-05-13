@echo off
setlocal
echo ===================================================
echo  CONFIGURANDO AMBIENTE PARA SCRIPT DE DIAGRAMA VENN
echo ===================================================

REM === Verifica se Python está instalado ===
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python não encontrado. Instalando...

    REM Baixa instalador do Python (Windows 64-bit embutido)
    powershell -Command "Invoke-WebRequest https://www.python.org/ftp/python/3.11.6/python-3.11.6-amd64.exe -OutFile python-installer.exe"

    echo Instalando Python silenciosamente...
    start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0

    REM Remove o instalador
    del python-installer.exe

    echo Python instalado com sucesso!
) else (
    echo Python já está instalado.
)

REM Atualiza pip
python -m ensurepip --upgrade
python -m pip install --upgrade pip

REM Instala bibliotecas necessárias
echo Instalando bibliotecas: pandas, matplotlib, venn, openpyxl
pip install pandas matplotlib venn openpyxl

echo.
echo ============================================
echo  AMBIENTE PRONTO!
echo  Para executar o script, digite:
echo     python venn_excel.py
echo ============================================
pause
