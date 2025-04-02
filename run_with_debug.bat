@echo off
chcp 65001 > nul
title PPT转视频工具 - 调试版

:: 简单的启动脚本，用于捕获错误信息
echo 正在以调试模式启动PPT转视频工具...
echo 如果程序崩溃，错误信息将被保存到 error.log 文件

:: 切换到脚本所在目录
cd /d "%~dp0"

:: 检查dist目录下是否有可执行文件
if exist "dist\PPT转视频工具.exe" (
    cd dist
    echo 已找到可执行文件，正在启动...
    
    :: 使用重定向捕获错误信息
    "PPT转视频工具.exe" 2> error.log
    
    :: 检查是否有错误
    if exist error.log (
        echo 程序执行过程中出现错误。
        echo 错误日志:
        type error.log
        echo.
        echo 按任意键退出...
        pause > nul
    )
) else (
    echo 未找到PPT转视频工具.exe
    echo 请先运行build.bat或fix_moviepy_build.py生成可执行文件
    echo.
    echo 按任意键退出...
    pause > nul
)
