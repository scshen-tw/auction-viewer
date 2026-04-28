@echo off
setlocal

cd /d "d:\VS Code\Auction"

set "PYTHON_EXE=C:\Program Files\Python311\python.exe"
set "DETAIL_LOG=auction_fetcher.log"
set "STATUS_LOG=daily_update_status.log"

echo ==== %date% %time% START daily update ====>> "%DETAIL_LOG%"
echo [%date% %time%] START daily update>> "%STATUS_LOG%"

"%PYTHON_EXE%" auction_fetcher.py --update >> "%DETAIL_LOG%" 2>&1
set "UPDATE_EXIT=%ERRORLEVEL%"

if "%UPDATE_EXIT%"=="0" (
    echo [%date% %time%] SUCCESS exit=%UPDATE_EXIT%>> "%STATUS_LOG%"
    echo ==== %date% %time% SUCCESS daily update ====>> "%DETAIL_LOG%"
    exit /b 0
)

echo [%date% %time%] FAILED exit=%UPDATE_EXIT% - see %DETAIL_LOG%>> "%STATUS_LOG%"
echo ==== %date% %time% FAILED daily update exit=%UPDATE_EXIT% ====>> "%DETAIL_LOG%"
exit /b %UPDATE_EXIT%
