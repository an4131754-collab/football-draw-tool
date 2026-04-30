@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PYTHON_EXE=C:\Users\USER\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

"%PYTHON_EXE%" "%SCRIPT_DIR%app.py"
endlocal
