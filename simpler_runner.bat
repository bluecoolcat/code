@echo off
chcp 65001 > nul
cd /d "%~dp0\dist"
if exist "PPT转视频工具.exe" (
  start "" "PPT转视频工具.exe"
) else (
  echo 错误：未找到可执行文件
  echo 请先运行build.bat生成程序
  pause
)
