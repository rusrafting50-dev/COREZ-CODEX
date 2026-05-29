cd /d "D:\Projects Claude Code\COREZ"
git merge --abort
git fetch origin
git checkout claude/main
git pull origin claude/main
pip install -r requirements.txt
python app.py
pause
