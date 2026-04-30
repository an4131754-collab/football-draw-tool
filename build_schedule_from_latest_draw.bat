@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PYTHON_EXE=C:\Users\USER\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "DRAW_FILE=%SCRIPT_DIR%outputs\latest\draw_result.json"

if not exist "%DRAW_FILE%" (
  echo 找不到 "%DRAW_FILE%"
  echo 請先執行抽籤網站並完成抽籤。
  pause
  exit /b 1
)

"%PYTHON_EXE%" "%SCRIPT_DIR%scheduler.py" --draw "%DRAW_FILE%"
pause
endlocal
