@echo off
setlocal

REM Prefer Python launcher with explicit 32-bit selector.
py -3-32 -c "import struct;print(8*struct.calcsize('P'))" >nul 2>&1
if %errorlevel%==0 (
  py -3-32 %*
  exit /b %errorlevel%
)

REM Fallback to explicit path if launcher is unavailable.
set "PY32=D:\ingenieur\py32\python.exe"
if not exist "%PY32%" (
  echo [ERR] 32-bit Python not found: %PY32%
  echo Please install 32-bit Python or update PY32 path in run_py32.bat
  exit /b 1
)

"%PY32%" %*
