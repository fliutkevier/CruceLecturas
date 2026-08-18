@echo off
REM Compila y despliega a la red en un solo paso.
REM Dejar este archivo junto a build.bat
cd /d "%~dp0"
call build.bat deploy
