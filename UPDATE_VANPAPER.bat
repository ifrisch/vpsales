@echo off
title Van Paper One-Click Update
color 0A
cd "C:\Users\Isaac\OneDrive - Van Paper Company\Python_Projects\Sales_Leaderboard"
"C:\Users\Isaac\AppData\Local\Programs\Python\Python313\python.exe" one_click_update.py
echo.
echo Pushing sync timestamp to live app...
git add last_sync.txt
git commit -m "Update sync timestamp from manual run"
git push
echo.
echo Live app will update in 1-2 minutes.
pause
