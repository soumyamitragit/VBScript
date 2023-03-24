@echo off
setlocal enabledelayedexpansion

REM Define array of prefixes to search for
set "prefixes=RELAY_PROPERTIES_ RELAY_ENV_"

REM Loop through prefixes and remove matching registry keys
for %%p in (%prefixes%) do (
    echo Removing registry keys starting with %%p...
    reg delete "HKCU\Environment\test" /f /va /k /s /se "%%p*" > nul
)

echo Done.
pause

REM