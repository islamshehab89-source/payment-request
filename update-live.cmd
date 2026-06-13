@echo off
title Publish Prices to Live Site
cd /d "%~dp0"

echo ===================================================
echo   Publishing your prices to the live site...
echo ===================================================
echo.

git add -A
git diff --cached --quiet
if %errorlevel%==0 (
  echo   Nothing changed since the last publish.
  echo   ^(Did you SAVE the Excel file first? Ctrl+S^)
  echo.
  pause
  exit /b 0
)

git commit -m "Update prices (%date% %time%)"
if errorlevel 1 goto :error

git push
if errorlevel 1 goto :error

echo.
echo   DONE  -  the live site refreshes in about 1-2 minutes:
echo   https://islamshehab89-source.github.io/payment-request/
echo.
pause
exit /b 0

:error
echo.
echo   Something went wrong. Take a screenshot of this window
echo   and send it to your developer.
echo.
pause
exit /b 1
