@echo off
echo ========================================
echo   Cleaning Python cache and restarting
echo ========================================

REM Kill all Python processes
taskkill /F /IM python.exe /T >nul 2>&1

REM Wait a moment
timeout /t 2 /nobreak >nul

REM Clean Python cache
echo Cleaning __pycache__ directories...
for /d /r . %%d in (__pycache__) do @if exist "%%d" rd /s /q "%%d"

REM Delete .pyc files
echo Cleaning .pyc files...
del /s /q *.pyc >nul 2>&1

echo.
echo Cache cleaned!
echo.
echo Starting server...
echo.

REM Start server
uv run python main.py
