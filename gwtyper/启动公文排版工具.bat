@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set PY=C:\Users\ASUS\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe
cd /d "%~dp0.."
start "" "%PY%" -m gwtyper.main
