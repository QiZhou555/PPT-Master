@echo off
REM PPT生成服务启动脚本
REM 此脚本自动设置必要的环境变量并启动Flask服务器

REM 设置环境变量
echo 设置环境变量...
set FLASK_SECRET_KEY=oeasy
set API_KEY=ms-4b457ae8-4cfd-4504-8ec2-8dc2fb930454
set API_BASE_URL=https://api-inference.modelscope.cn/v1
set FLASK_HOST=0.0.0.0
set FLASK_PORT=5001

REM 检查Python是否安装
python --version >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo 错误: 未找到Python。请确保已安装Python并添加到系统PATH中。
    pause
    exit /b 1
)

REM 检查依赖是否已安装
echo 检查必要的依赖...
pip list | findstr "flask" >nul
if %ERRORLEVEL% neq 0 (
    echo 警告: 未找到Flask。正在尝试安装必要的依赖...
    pip install -r requirements.txt
    if %ERRORLEVEL% neq 0 (
        echo 错误: 安装依赖失败。请手动运行: pip install -r requirements.txt
        pause
        exit /b 1
    )
)

REM 显示配置信息
echo ===================================================
echo PPT生成服务启动配置
echo 服务地址: http://0.0.0.0:5001
echo 本地访问: http://127.0.0.1:5001
echo 密钥: oeasy
echo ===================================================
echo 正在启动Flask服务器...
echo 按 Ctrl+C 停止服务

REM 启动服务器
python app.py