@echo off
echo Running ATT Export Script...
echo.

REM Set parameters
set SOURCE_MDB=%~1
set TARGET_HOST=%~2
set EXPORT_DATE=%~3

REM Display parameters
echo Source MDB: %SOURCE_MDB%
echo Target Host: %TARGET_HOST%
if not "%EXPORT_DATE%"=="" (
    echo Export Date: %EXPORT_DATE%
)
echo.

REM Check if compiled JavaScript exists
if not exist "dist\attexport.js" (
    echo Error: dist\attexport.js not found
    echo Please compile the TypeScript file first using: npm run build
    pause
    exit /b 1
)

REM Run the script
echo Starting export...
node dist\attexport.js "%SOURCE_MDB%" "%TARGET_HOST%" %EXPORT_DATE%

echo.
echo Export completed!
pause
