@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
"%SCRIPT_DIR%\.venv\Scripts\pythonw.exe" "%SCRIPT_DIR%clipboard_manager.py"
