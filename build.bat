@echo off
echo ============================================================
echo  Verificando Python...
echo ============================================================
python --version
if errorlevel 1 (
    echo.
    echo [ERRO] Python nao foi encontrado no PATH.
    echo Instale o Python em https://www.python.org/downloads/
    echo e marque a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Verificando arquivos necessarios...
echo ============================================================
if not exist "index.html" (
    echo [ERRO] index.html nao encontrado nesta pasta.
    echo Todos os arquivos precisam estar juntos, na mesma pasta:
    echo app.py, index.html, config.py, pdf_reader.py, excel_writer.py
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Instalando dependencias...
echo ============================================================
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao instalar as dependencias do requirements.txt.
    echo Veja a mensagem de erro acima.
    pause
    exit /b 1
)

python -m pip install pyinstaller
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao instalar o pyinstaller.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Gerando executavel (modo pasta - mais estavel)...
echo ============================================================
python -m PyInstaller --noconsole --onedir --name ImportadorBorderos --add-data "index.html;." app.py
if errorlevel 1 (
    echo.
    echo [ERRO] O PyInstaller falhou. Veja a mensagem de erro acima
    echo e me envie o texto para eu te ajudar a corrigir.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Pronto! Va ate a pasta: dist\ImportadorBorderos
echo  O executavel e o arquivo: ImportadorBorderos.exe
echo  IMPORTANTE: distribua a PASTA "ImportadorBorderos" inteira,
echo  nao so o .exe sozinho.
echo ============================================================
pause
