@echo off
echo 请选择使用说明书版本：
echo 1. 完整版（详细说明）
echo 2. 极简版（快速上手）
set /p choice=请输入选项 (1 或 2):

if "%choice%"=="1" (
    start 使用说明书.html
) else if "%choice%"=="2" (
    start 使用说明书-极简版.html
) else (
    echo 无效选项，默认打开完整版
    start 使用说明书.html
)