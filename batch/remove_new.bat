@echo off
setlocal

set "Action=Find"
set "RegKey=HKEY_CURRENT_USER\Environment\Test"
set "Search=RELAY_"

REM Check if Specified Registry Exists
%SystemRoot%\System32\reg.exe query %RegKey% 1>nul 2>nul
if not errorlevel 1 goto RunSearch

if /I "%Action%"=="Delete" (
    echo Deleted %DeleteCounter% key%DeletePlural% of %FoundCounter% key%FoundPlural% containing "%Search%".
) else (
    echo Found %FoundCounter% key%FoundPlural% containing "%Search%".
)


:EndBatch
endlocal
echo.
echo Exit with any key ...
pause >nul

:EndBatch