@echo off
setlocal

REM Build .exe para o app Busca NFSe
REM Uso: execute este .bat no Windows (CMD) na raiz do projeto.

cd /d "%~dp0"

echo [1/6] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
  echo Python nao encontrado no PATH.
  exit /b 1
)

echo [2/6] Criando/ativando venv...
if not exist .venv (
  python -m venv .venv
)
call .venv\Scripts\activate.bat
if errorlevel 1 exit /b 1

echo [3/6] Instalando dependencias de build...
python -m pip install --upgrade pip
python -m pip install pyinstaller customtkinter pandas openpyxl requests
if errorlevel 1 exit /b 1

echo [4/6] Gerando icone...
python scripts\generate_icon.py
if errorlevel 1 exit /b 1

echo [5/6] Gerando executavel...
pyinstaller --noconfirm --clean app_nfse_lote_excel.spec
if errorlevel 1 exit /b 1

echo [6/6] Copiando config base para dist...
if exist config.json copy /Y config.json dist\BuscaNFSe\config.json >nul

echo.
echo Build finalizado com sucesso!
echo Executavel: dist\BuscaNFSe\BuscaNFSe.exe
endlocal
