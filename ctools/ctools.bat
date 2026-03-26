@echo off
chcp 65001 >nul
echo 你好，欢迎使用 ctools

:: 使用 PowerShell 运行以获得更好的 UTF-8 支持
powershell -Command "& { $env:DOTNET_SYSTEM_CONSOLE_ALLOW_ANSI_COLOR_REDIRECTION='1'; [Console]::InputEncoding = [System.Text.Encoding]::UTF8; [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; dotnet run --project 'E:\cqh\code\my_c#\ctools\ctool.csproj' --no-build }"

echo 你好，欢迎使用 ctools