@echo off
cd /d "d:\VS Code\Auction"

"C:\Program Files\Python311\python.exe" auction_fetcher.py --update >> auction_fetcher.log 2>&1

git add auction_viewer.html auction_stocks.json auction_cbs.json price_cache.json
git diff --cached --quiet && goto :no_changes

for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set TODAY=%%a/%%b/%%c
for /f "tokens=1-2 delims=: " %%a in ("%time%") do set NOW=%%a:%%b
git commit -m "自動更新資料 %TODAY% %NOW%"
git push >> auction_fetcher.log 2>&1
goto :end

:no_changes
echo %date% %time% - 無新資料，略過 push >> auction_fetcher.log

:end
