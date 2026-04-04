@echo off
echo ===================================
echo 正在注册 SolidWorks 插件...
echo ===================================

:: 以管理员身份运行 regasm
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe "%~dp0bin\Debug\net48\plugin.dll"  /codebase

pause
