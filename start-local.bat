@echo off
setlocal
cd /d "%~dp0"

where netlify >nul 2>&1
if %errorlevel% neq 0 (
  echo Netlify CLI not found. Installing...
  npm install -g netlify-cli
  if %errorlevel% neq 0 (
    echo Failed to install Netlify CLI. Please run as Administrator:
    echo npm install -g netlify-cli
    pause
    exit /b 1
  )
)

echo Starting local app with Netlify Dev...
echo Open: http://localhost:8888
netlify dev

endlocal
