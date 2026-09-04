@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py") else (set "PY=python")
%PY% -m pip install pyinstaller
if %errorlevel% neq 0 exit /b 1
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name "Copy-Paste-Manager" copy_paste_manager.py
if %errorlevel% neq 0 exit /b 1
echo.
echo SUCCESS: dist\Copy-Paste-Manager.exe
pause
