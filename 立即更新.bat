@echo off
rem Macro Dashboard - Manual update (double-click to run)
rem Check update.log in the same folder for detailed results.
title Macro Dashboard Update
echo ============================================
echo  Updating macro indicators, please wait...
echo  (see update.log for details)
echo ============================================
echo.
call "%~dp0run_update.bat"
echo.
echo ============================================
echo  Done. The dashboard has been refreshed.
echo ============================================
echo.
pause
