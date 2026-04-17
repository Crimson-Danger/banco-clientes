@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Ambiente virtual nao encontrado. Crie com:
  echo python -m venv .venv
  pause
  exit /b 1
)
call ".venv\Scripts\activate.bat"
python run_inss_app.py
