@echo off
setlocal
where py >nul 2>nul
if errorlevel 1 (
  echo Не найден Python Launcher. Установите Python 3.11 или 3.12 с python.org.
  exit /b 1
)
if not exist .venv\Scripts\python.exe py -3 -m venv .venv
.venv\Scripts\python.exe -m pip --version >nul 2>nul
if errorlevel 1 (
  echo Виртуальная среда неполная. Пересоздаю .venv...
  py -3 -m venv --clear .venv
)
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo Установка завершена. Запустите run_dashboard.bat
