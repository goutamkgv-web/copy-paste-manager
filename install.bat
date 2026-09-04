@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py"
) else (
  where python >nul 2>nul
  if %errorlevel% neq 0 (
    echo Python 3.10+ was not found.
    pause
    exit /b 1
  )
  set "PY=python"
)
%PY% -m pip install -r requirements.txt
if %errorlevel% neq 0 (
  echo Installation failed.
  pause
  exit /b 1
)
echo Setup complete. Run run.bat.
pause
