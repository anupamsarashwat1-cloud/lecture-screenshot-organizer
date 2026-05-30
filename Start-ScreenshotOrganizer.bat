@echo off
title Lecture Screenshot Organizer
echo ==========================================================
echo Starting Lecture Screenshot Organizer Agent...
echo ==========================================================
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Start-ScreenshotOrganizer.ps1"
pause
