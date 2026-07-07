@echo off
setlocal
cd /d %TEMP%
pushd "%~dp0"
if errorlevel 1 (
  echo pushd failed: %~dp0
  exit /b 1
)
echo PWD: %CD%
echo.
echo === clasp push ===
call clasp.cmd push --force
if errorlevel 1 (
  popd
  exit /b 1
)
echo.
echo === clasp deploy ===
call clasp.cmd deploy -i AKfycbxJoWYd4xk6Tycy6mLX0le-2sVwlJP6qID-HRBbMOJ5AI8fvr83HDpUJiKyOq38_g3D0Q -d "DOMAIN公開"
set ERR=%ERRORLEVEL%
popd
exit /b %ERR%
