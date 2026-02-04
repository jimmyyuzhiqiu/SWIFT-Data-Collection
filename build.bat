@echo off
REM SWIFT Data Collection - 打包脚本
REM 在 Windows 上运行此脚本生成 exe

echo ========================================
echo SWIFT Data Collection - 打包工具
echo ========================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未找到 Python，请先安装 Python 3.9+
    pause
    exit /b 1
)

echo ✅ Python 已安装
echo.

REM 安装依赖
echo 📦 安装依赖中...
pip install -r requirements.txt
if errorlevel 1 (
    echo ❌ 依赖安装失败
    pause
    exit /b 1
)

echo ✅ 依赖安装完成
echo.

REM 运行打包脚本
echo 🔨 开始打包...
python build.py
if errorlevel 1 (
    echo ❌ 打包失败
    pause
    exit /b 1
)

echo.
echo ✅ 打包完成！
echo 📁 输出文件: dist\SWIFT_Data_Collection.exe
echo.
pause
