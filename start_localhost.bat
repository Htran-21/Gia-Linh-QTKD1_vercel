@echo off
chcp 65001 >nul
echo ============================================================================
echo   KHOI DONG EDTECH FUTURE AI AGENT SERVER TREN LOCALHOST (PORT 8080)
echo   Truy cap trinh duyet tai: http://localhost:8080
echo ============================================================================
start http://localhost:8080
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1" -Port 8080
pause
