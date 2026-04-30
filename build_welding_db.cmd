@echo off
setlocal EnableExtensions

rem This CMD is ASCII-only on purpose.
rem Put this file in the same folder as build_welding_db.py and Statistika xlsm.
rem It works from a UNC network share because pushd maps the share to a temporary drive letter.

pushd "%~dp0"
if errorlevel 1 (
    echo ERROR: cannot open script folder:
    echo %~dp0
    pause
    exit /b 1
)

if not exist "build_welding_db.py" (
    echo ERROR: build_welding_db.py not found in:
    cd
    popd
    pause
    exit /b 1
)

echo.
echo Building welding_db.json and welding_db.json.gz...
echo Folder:
cd
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "build_welding_db.py" "*.xlsm" -o "welding_db.json"
) else (
    python "build_welding_db.py" "*.xlsm" -o "welding_db.json"
)

set "ERR=%ERRORLEVEL%"
echo.

if "%ERR%"=="0" (
    echo DONE.
    echo Created or updated:
    echo welding_db.json
    echo welding_db.json.gz
) else (
    echo ERROR: build failed. Exit code: %ERR%
)

popd
echo.
pause
exit /b %ERR%
