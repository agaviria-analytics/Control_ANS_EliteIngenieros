@echo off
cd /d "%~dp0"
call venv\Scripts\activate
python menu_proyecto_ans.py
pause
