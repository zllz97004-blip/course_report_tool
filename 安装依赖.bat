@echo off
chcp 65001 >nul
cd /d %~dp0

echo ========================================
echo course_report_tool 依赖安装
echo ========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo 未找到 python，请先安装 Python 3.8-3.10，并勾选 Add Python to PATH。
    pause
    exit /b 1
)

if not exist .venv (
    echo 正在创建本地虚拟环境 .venv ...
    python -m venv .venv
    if errorlevel 1 (
        echo 创建虚拟环境失败。
        pause
        exit /b 1
    )
)

echo 正在升级 pip ...
.venv\Scripts\python.exe -m pip install --upgrade pip -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com

echo.
echo 正在安装项目依赖 ...
.venv\Scripts\python.exe -m pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com
if errorlevel 1 (
    echo.
    echo 依赖安装失败。可尝试更换网络，或将上方镜像地址改为清华源：
    echo https://pypi.tuna.tsinghua.edu.cn/simple
    pause
    exit /b 1
)

echo.
echo 依赖安装完成。
pause
