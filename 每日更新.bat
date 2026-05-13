@echo off
setlocal

cd /d "d:\vscode\Auction"

set "PYTHON_EXE=C:\Program Files\Python311\python.exe"
set "DETAIL_LOG=auction_fetcher.log"
set "RUNNER_LOG=daily_update_runner.log"
set "STATUS_LOG=daily_update_status.log"
set "GIT_TERMINAL_PROMPT=0"
set "GCM_INTERACTIVE=never"
if exist "C:\Users\User\.ssh\id_ed25519_auction_viewer" set "GIT_SSH_COMMAND=ssh -i C:/Users/User/.ssh/id_ed25519_auction_viewer -o IdentitiesOnly=yes -o StrictHostKeyChecking=accept-new"

echo ==== %date% %time% START daily update ====>> "%DETAIL_LOG%"
echo [%date% %time%] START daily update>> "%STATUS_LOG%"

"%PYTHON_EXE%" auction_fetcher.py --update >> "%RUNNER_LOG%" 2>&1
set "UPDATE_EXIT=%ERRORLEVEL%"

if "%UPDATE_EXIT%"=="0" (
    echo [%date% %time%] SUCCESS exit=%UPDATE_EXIT% >> "%STATUS_LOG%"
    echo ==== %date% %time% SUCCESS daily update ====>> "%DETAIL_LOG%"
    exit /b 0
)

echo [%date% %time%] FAILED exit=%UPDATE_EXIT% - see %DETAIL_LOG% >> "%STATUS_LOG%"
echo ==== %date% %time% FAILED daily update exit=%UPDATE_EXIT% ====>> "%DETAIL_LOG%"
exit /b %UPDATE_EXIT%
