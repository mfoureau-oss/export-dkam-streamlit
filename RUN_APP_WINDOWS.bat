@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [1/2] Creation de l'environnement Python (.venv)...
  python -m venv .venv
  if errorlevel 1 (
    echo ERREUR: Python n'est pas installe ou n'est pas dans le PATH.
    pause
    exit /b 1
  )
)

echo [2/2] Installation/MAJ des dependances...
call ".venv\Scripts\activate.bat"
pip install -r requirements.txt

echo Lancement de l'application...
streamlit run app\streamlit_app.py
