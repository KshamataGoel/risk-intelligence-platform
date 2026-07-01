@echo off
echo ================================================
echo   Risk Intelligence Platform - Starting Up
echo ================================================
echo.

REM Check PATH first, then fallback to known local install
where node >nul 2>&1
if %errorlevel% equ 0 (
    set NODE_EXE=node
) else if exist "%USERPROFILE%\node\node.exe" (
    set NODE_EXE=%USERPROFILE%\node\node.exe
    set NPM_EXE=%USERPROFILE%\node\npm.cmd
) else (
    echo ERROR: Node.js not found.
    echo Please install Node.js from https://nodejs.org
    echo Or place node.exe in %USERPROFILE%\node\
    pause
    exit /b 1
)

if not exist ".env" (
    if exist ".env.example" (
        echo Creating .env from template...
        copy .env.example .env >nul
        echo IMPORTANT: Edit .env and add your GROQ_API_KEY
        echo.
    )
)

echo Installing dependencies...
if defined NPM_EXE (
    call "%NPM_EXE%" install --silent
) else (
    call npm install --silent
)

echo.
echo Starting server...
echo Open http://localhost:3000 in your browser
echo.
"%NODE_EXE%" server.js
pause
