@echo off
chcp 65001 >nul
cd /d "D:\Projects Claude Code\COREZ"

if "%~1"=="--relaunched" goto :run

echo [1/4] Останавливаем старый Flask (порт 5000)...
for /f "tokens=5" %%p in ('netstat -ano ^| findstr :5000 ^| findstr LISTENING') do taskkill /f /pid %%p 1>nul 2>nul
timeout /t 2 /nobreak >nul

echo [2/4] Получаем обновления...
git pull origin main

rem git pull мог изменить сам этот .bat-файл — перезапускаемся новым
rem процессом, чтобы дочитать файл с диска заново, а не со старого
rem (теперь невалидного) смещения в этом же процессе cmd.exe.
start "" "%~f0" --relaunched
exit /b

:run
echo [3/4] Проверяем зависимости...
pip install -r requirements.txt --quiet

echo [4/4] Запускаем Flask...
python app.py
pause
