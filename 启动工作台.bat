@echo off
chcp 65001 >nul
cd /d %~dp0

if not exist .venv\Scripts\python.exe (
    echo 未找到本地虚拟环境 .venv。
    echo 请先双击运行 “安装依赖.bat”。
    pause
    exit /b 1
)

.venv\Scripts\python.exe -m streamlit run app.py
pause
