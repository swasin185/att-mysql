@echo off
echo Starting Node.js 24 installation and build process...
echo.

echo Copying source files...
copy ..\attexport.ts .
echo Source files copied.

echo Installing dependencies...
call npm run install
echo Dependencies installed.

echo Building TypeScript project...
call npm run build
echo Build completed.

echo Starting the application...
call npm run serve
echo Application started.

pause
