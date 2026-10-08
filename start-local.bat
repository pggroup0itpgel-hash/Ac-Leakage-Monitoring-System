@echo off
setlocal
cd /d "%~dp0"

echo Starting local AC Leakage Monitoring System...
echo Open: http://localhost:8888
node server.js

endlocal
