# 打包为 GUI exe + CLI exe（需先激活 .venv 并安装依赖）
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
.\.venv\Scripts\pyinstaller.exe build.spec --noconfirm --clean
Write-Host ""
Write-Host "完成："
Write-Host "  dist\文档转Markdown.exe  （图形界面）"
Write-Host "  dist\doc2md-cli.exe      （命令行 / stdin）"
