@echo off
chcp 65001 >nul
cd /d "D:\Projects Claude Code\COREZ"

echo [1/4] Останавливаем старый Flask...
taskkill /f /im python.exe 1>nul 2>nul
timeout /t 2 /nobreak >nul

echo [2/4] Получаем обновления...
git pull origin main

echo [3/4] Проверяем зависимости...
pip install -r requirements.txt --quiet

echo [4/4] Запускаем Flask...
python app.py
pause
