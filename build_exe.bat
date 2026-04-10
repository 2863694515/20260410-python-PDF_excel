@echo off
setlocal
cd /d %~dp0

echo [1/3] 安装打包依赖...
py -m pip install --upgrade pip pyinstaller >nul
if errorlevel 1 (
  echo 安装 PyInstaller 失败，请检查 Python 环境。
  exit /b 1
)

echo [2/3] 清理旧构建...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist app.spec del /q app.spec

echo [3/3] 打包 exe...
py -m PyInstaller --noconfirm --onefile --name PDF2Excel-WebUI --add-data "templates;templates" --add-data "static;static" app.py
if errorlevel 1 (
  echo 打包失败。
  exit /b 1
)

echo.
echo 打包成功：dist\PDF2Excel-WebUI.exe
echo 直接把这个 exe 发给别人即可使用。
endlocal
