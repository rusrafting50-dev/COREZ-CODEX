cd /d "D:\Projects Claude Code\COREZ"
git merge --abort
git fetch origin
git checkout claude/corez-rafting-system-rxNfz
git pull origin claude/corez-rafting-system-rxNfz
pip install -r requirements.txt
python app.py
pause
