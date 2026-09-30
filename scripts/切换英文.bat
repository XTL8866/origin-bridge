@echo off
setlocal enabledelayedexpansion
title 切换 Origin 界面为英文

echo ========================================
echo   将 Origin 界面切换为【英文】
echo ========================================
echo   说明：Origin 没有图形化语言开关，
echo         界面语言由注册表值 Language 控制。
echo.

echo [1/4] 关闭正在运行的 Origin ...
taskkill /IM Origin64.exe /F >nul 2>nul
taskkill /IM Origin.exe /F >nul 2>nul
timeout /t 2 /nobreak >nul

echo [2/4] 定位 Origin 注册表键 ...
set "KEY="
for /f "tokens=*" %%K in ('reg query "HKCU\Software\OriginLab" 2^>nul ^| findstr /i "Origin"') do set "KEY=%%K"
if not defined KEY (
    echo     [失败] 未找到 Origin 安装记录
    echo     请确认 Origin 已安装，并至少手动运行过一次
    echo.
    pause
    exit /b 1
)
set "KEY=!KEY:HKEY_CURRENT_USER=HKCU!"
echo     找到键：!KEY!

echo [3/4] 写入 Language = E ...
reg add "!KEY!" /v Language /t REG_SZ /d E /f
if errorlevel 1 (
    echo     [失败] 注册表写入失败
    echo     请右键本文件选择“以管理员身份运行”后重试
    echo.
    pause
    exit /b 1
)

echo [4/4] 校验写入结果 ...
reg query "!KEY!" /v Language

echo.
echo ========================================
echo   完成！请重新打开 Origin 查看英文界面。
echo.
echo   备注：Origin 的中文语言包非 100%% 覆盖，
echo         部分高级分析工具与第三方 App 仍为英文，属正常现象。
echo ========================================
echo.
pause
