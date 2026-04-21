cd /d "D:\Projects Claude Code\COREZ"
git merge --abort
git fetch origin
git checkout claude/corez-rafting-system-enABl
git pull origin claude/corez-rafting-system-enABl
pip install -r requirements.txt
python app.py
pause
