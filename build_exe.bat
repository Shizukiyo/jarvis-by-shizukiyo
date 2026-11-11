@echo off
cd /d "%~dp0"
call .venv\Scripts\activate
python -m pip install pyinstaller
pyinstaller --noconfirm --onedir --name Jarvis --hidden-import win32com.client jarvis.py
echo.
echo Готово: dist\Jarvis\Jarvis.exe
pause
