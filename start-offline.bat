@echo off
title LLM Gateway for Excel - Local Offline Server
cd /d "%~dp0"
echo ===================================================
echo   LLM Gateway for Excel - Local Offline Server
echo ===================================================
echo Status: Running local server on https://localhost:3141
echo You can now open Microsoft Excel and launch the Add-in!
echo Press Ctrl+C or close this window to stop the server.
echo ===================================================
npm run dev
