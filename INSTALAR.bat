@echo off
echo Instalando dependencias da Pesquisa de Precos...
python -m pip install --upgrade pip
python -m pip install openpyxl playwright
python -m playwright install chromium
if not exist "%~dp0lojas.json" copy "%~dp0lojas.exemplo.json" "%~dp0lojas.json"
echo @echo off> "%~dp0EXECUTAR.bat"
echo python "%%~dp0atualizar_precos.py">> "%~dp0EXECUTAR.bat"
echo.
echo Instalacao concluida. Edite lojas.json e use EXECUTAR.bat para abrir o programa.
pause
