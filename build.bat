@echo off
REM Builds dist\Cobranzas.exe. Run on Windows with Python 3.8 installed.
cd /d "%~dp0"

if not exist .venv (
    py -3.8 -m venv .venv 2>nul || python -m venv .venv
)
call .venv\Scripts\activate.bat
python -m pip install -r requirements.txt || goto :error
python -m unittest discover -s tests || goto :error
python -m PyInstaller --noconfirm --onefile --windowed --name Cobranzas --collect-data certifi app.py || goto :error

if not exist dist\credenciales.ini copy /Y credenciales.example.ini dist\credenciales.ini >nul
if exist logo.png copy /Y logo.png dist\ >nul
echo.
echo Listo: dist\Cobranzas.exe
echo Copie Cobranzas.exe y credenciales.ini (completado) a la PC con Windows 7.
pause
exit /b 0

:error
echo.
echo ERROR: la compilacion fallo. Revise los mensajes de arriba.
pause
exit /b 1
