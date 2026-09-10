@echo off
setlocal
if not exist "%~dp0node_modules\npm\bin\npm-cli.js" goto missing
node "%~dp0node_modules\npm\bin\npm-cli.js" %*
exit /b %errorlevel%
:missing
echo Local npm is missing. Install dependencies using a Node.js distribution with npm, then run npm ci. 1>&2
exit /b 1
