@echo off
cd /d "%~dp0"
echo Mise a jour depuis GitHub...
git pull --ff-only
echo.
python extract_projets_engineering.py
