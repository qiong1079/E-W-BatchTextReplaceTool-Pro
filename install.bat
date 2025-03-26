@echo off
echo ========================================
echo          文档批量替换工具安装
echo ========================================
echo.

echo 检查 Python 安装...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到 Python。请先安装 Python 3.7 或更高版本。
    echo 您可以从 https://www.python.org/downloads/ 下载安装。
    pause
    exit /b 1
)

echo Python 已安装, 开始安装必要的库...
echo.
echo 安装 pywin32...
pip install pywin32
if %errorlevel% neq 0 (
    echo 安装过程中出现错误。请确保网络连接正常，或尝试以管理员身份运行此脚本。
    pause
    exit /b 1
)

echo.
echo ========================================
echo            安装成功完成!
echo ========================================
echo.
echo 您现在可以通过执行以下步骤使用文档批量替换工具:
echo 1. 编辑 config.py 文件设置替换规则和文件路径
echo 2. 运行 docx_formatted_replace.py 开始处理文件
echo.
echo 更多信息请查看"使用说明.txt"文件。
echo.
pause 