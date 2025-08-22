@echo off
set "exePath=%~dp0AudioPortal.exe"
set "startupFolder=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "shortcutName=AudioPortal.lnk"

:: 检查 AudioPortal.exe 是否存在
if not exist "%exePath%" (
    echo AudioPortal.exe not found at "%exePath%".
    exit /b
)

:: 创建快捷方式
powershell -Command "$s=(New-Object -COMObject WScript.Shell).CreateShortcut('%startupFolder%\%shortcutName%'); $s.TargetPath='%exePath%'; $s.WorkingDirectory='%~dp0'; $s.Save()"

if %errorlevel% neq 0 (
    echo Failed to create shortcut. Please check permissions.
) else (
    echo AudioPortal.exe has been added to startup.
)

pause