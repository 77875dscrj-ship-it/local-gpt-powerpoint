@echo off
setlocal
set "ROOT=%~dp0"
set "NODE=C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
if not exist "%NODE%" set "NODE=node"
cd /d "%ROOT%"
"%NODE%" "%ROOT%server.js"
