@echo off
echo Running conversion script...
:: This runs the script in your normal user space matching your Office install
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0The-Archivist.ps1"
