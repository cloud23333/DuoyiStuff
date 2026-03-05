@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%" >nul 2>&1 || (
  echo [ERROR] Failed to enter project directory: "%SCRIPT_DIR%"
  exit /b 1
)

set "PY_CMD="
where py >nul 2>&1
if not errorlevel 1 (
  py -3.13 -c "import sys" >nul 2>&1 && set "PY_CMD=py -3.13"
  if not defined PY_CMD (
    py -3 -c "import sys" >nul 2>&1 && set "PY_CMD=py -3"
  )
)
if not defined PY_CMD (
  where python >nul 2>&1 && set "PY_CMD=python"
)
if not defined PY_CMD (
  echo [ERROR] Python was not found in PATH.
  goto :fail
)

set "SPEC_FILE="
for %%F in ("%SCRIPT_DIR%*.spec") do (
  set "SPEC_FILE=%%~fF"
  goto :spec_found
)

echo [ERROR] No .spec file found in "%SCRIPT_DIR%"
goto :fail

:spec_found
echo [INFO] Python command: %PY_CMD%
echo [INFO] Spec file: "%SPEC_FILE%"

echo [STEP] Installing/Updating build dependencies...
call %PY_CMD% -m pip install -U pyinstaller PyQt6
if errorlevel 1 (
  echo [ERROR] Dependency installation failed.
  goto :fail
)

if /I "%~1"=="clean" (
  echo [STEP] Cleaning build and dist folders...
  if exist build rmdir /s /q build
  if exist dist rmdir /s /q dist
)

echo [STEP] Running PyInstaller...
call %PY_CMD% -m PyInstaller --noconfirm "%SPEC_FILE%"
if errorlevel 1 (
  echo [ERROR] PyInstaller build failed.
  goto :fail
)

echo [DONE] Build completed.
if exist dist (
  echo [INFO] Output folder: "%SCRIPT_DIR%dist"
  echo [INFO] EXE files:
  dir /b /a:-d "%SCRIPT_DIR%dist\*.exe"
) else (
  echo [WARN] dist folder was not created.
)

popd >nul
exit /b 0

:fail
popd >nul
exit /b 1
