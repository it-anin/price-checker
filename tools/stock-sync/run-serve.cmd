@echo off
cd /d "%~dp0"
"C:\Program Files\nodejs\node.exe" --env-file=.env sync.mjs --serve >> serve.log 2>&1
