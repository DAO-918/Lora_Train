@echo off
setlocal enabledelayedexpansion

:: 设置编码为UTF-8
chcp 65001 > nul

:: 获取当前bat文件所在目录
set "BAT_DIR=%~dp0"

:: 获取Lora_0_Start.py脚本的路径
set "SCRIPT_PATH=D:\Code\MY_ComfyUI\# 模型训练\#Lora_0_Start.py"

:: 如果SCRIPT_PATH不存在，尝试使用相对路径
if not exist "!SCRIPT_PATH!" (
    set "SCRIPT_PATH=%~dp0#Lora_0_Start.py"
)

:: 检查脚本是否存在
if not exist "!SCRIPT_PATH!" (
    echo 错误: 无法找到#Lora_0_Start.py脚本
    echo 请确保脚本位于以下位置之一:
    echo 1. D:\Code\MY_ComfyUI\# 模型训练\#Lora_0_Start.py
    echo 2. %~dp0Lora_0_Start.py
    pause
    exit /b 1
)

echo 正在运行Lora训练初始化脚本...
echo 使用当前目录作为项目路径: %BAT_DIR%

:: 运行Python脚本，使用--bat_dir参数指示使用bat文件所在目录作为项目路径
python "!SCRIPT_PATH!" --bat_dir

:: 检查脚本执行结果
if %errorlevel% neq 0 (
    echo 脚本执行失败，错误代码: %errorlevel%
) else (
    echo 脚本执行成功
)

pause