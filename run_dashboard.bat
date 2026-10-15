@echo off
setlocal
if not exist .venv\Scripts\python.exe (
  echo Сначала запустите install_windows.bat
  exit /b 1
)
call .venv\Scripts\activate.bat
python -m streamlit run app.py
