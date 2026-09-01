@echo off
set PYTHONIOENCODING=utf-8
chcp 65001 > nul
cd /d "%~dp0"
start "" pythonw -m docconv --gui
