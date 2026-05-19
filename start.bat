@echo off
echo ================================================
echo   Risk Intelligence Platform - Starting Up
echo ================================================
echo.

where node >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Node.js is not installed.
    echo Please install Node.js from https://nodejs.org
    pause
    exit /b 1
)

if not exist ".env" (
    echo Creating .env from template...
    copy .env.example .env >nul
    echo IMPORTANT: Edit .env and add your GROQ_API_KEY
    echo.
)

echo Installing dependencies...
call npm install --silent
if %errorlevel% neq 0 (
    echo ERROR: npm install failed.
    pause
    exit /b 1
)

echo.
echo Starting server...
echo Open http://localhost:3000 in your browser
echo.
node server.js
pause
