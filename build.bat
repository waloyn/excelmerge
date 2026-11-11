@echo off
chcp 65001 >nul
echo 正在切换 Go 1.20 工具链 (保证 Win7 兼容)...
call g use 1.20
echo 正在编译 Windows 原生资源 (Manifest Comctl32 v6)...
call "%GOPATH%\bin\rsrc.exe" -manifest manifest.xml -o rsrc.syso
echo 正在构建绿色单文件 EXE (无控制台黑框)...
go build -ldflags="-H windowsgui -s -w" -o ExcelMerge.exe .
if %ERRORLEVEL% EQU 0 (
    echo 构建成功: ExcelMerge.exe
) else (
    echo 构建失败，请检查报错！
)
pause
