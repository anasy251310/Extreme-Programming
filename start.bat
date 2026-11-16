@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv" (
    echo [1/3] 创建虚拟环境...
    python -m venv .venv
)

echo [2/3] 检查依赖...
call .venv\Scripts\activate.bat
python -m pip install -q --disable-pip-version-check -r requirements.txt

echo [3/3] 启动通讯录： http://127.0.0.1:5000
python app.py
pause
