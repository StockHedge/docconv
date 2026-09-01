@echo off
set PYTHONIOENCODING=utf-8
chcp 65001 > nul
cd /d "%~dp0"
echo 필요한 패키지를 설치합니다...
python -m pip install -r requirements.txt
echo.
echo 설치가 끝났습니다. 환경을 확인합니다.
python -m docconv --doctor
pause
