@echo off
REM ─────────────────────────────────────────────────────────
REM  docx2pptx — launcher (Windows)
REM  Double-click this file to start
REM ─────────────────────────────────────────────────────────

cd /d "%~dp0"

IF NOT EXIST node_modules (
  echo Installing dependencies (first run only)...
  npm install pptxgenjs react react-dom react-icons sharp
)

echo.
echo Starting docx2pptx server...
echo Open your browser at: http://localhost:4242
echo Press Ctrl+C to stop.
echo.

node server.js
pause
