@echo off
setlocal

REM Prefer Python launcher with explicit 32-bit selector.
py -3-32 -c "import struct;print(8*struct.calcsize('P'))" >nul 2>&1
if %errorlevel%==0 (
  py -3-32 %*
  exit /b %errorlevel%
)

REM Fallback: use a user-provided 32-bit Python path.
REM Set PY32 to your 32-bit python.exe (system/user env var or edit below).
if defined PY32 (
  if exist "%PY32%" (
    "%PY32%" %*
    exit /b %errorlevel%
  )
  echo [ERR] 32-bit Python not found: %PY32%
  echo Please fix PY32 or install 32-bit Python.
  exit /b 1
)

echo [ERR] 32-bit Python not found.
echo Please set PY32 to your 32-bit python.exe or install Python 32-bit.
exit /b 1
