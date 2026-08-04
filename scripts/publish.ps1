# 一键发布到 GitHub（需先完成登录，见下方说明）
# 用法：在仓库根目录执行 .\scripts\publish.ps1

$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)

$version = "v1.7.2"
Write-Host "=== 文档转 Markdown 发布脚本 ===" -ForegroundColor Cyan
Write-Host "版本: $version"
Write-Host ""

# 1. 检查 gh
$gh = Get-Command gh -ErrorAction SilentlyContinue
if (-not $gh) {
    Write-Host "未检测到 GitHub CLI (gh)。" -ForegroundColor Yellow
    Write-Host "请先安装: winget install GitHub.cli"
    Write-Host "然后登录: gh auth login"
    exit 1
}

# 2. 检查登录
gh auth status 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "请先登录 GitHub: gh auth login" -ForegroundColor Yellow
    exit 1
}

# 3. 创建仓库（若不存在）
$repo = "wluser3362203440/doc2md"
gh repo view $repo 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "创建仓库 $repo ..."
    gh repo create $repo --public --source=. --remote=origin --description "文档转 Markdown — Windows 桌面工具 v1.7.2"
}

# 4. 推送代码
Write-Host "推送代码..."
git push -u origin main
git tag -a $version -m "Release $version" -f
git push origin $version --force

# 5. 打包 exe（若 dist 不存在）
if (-not (Test-Path "dist\doc2md-cli.exe")) {
    Write-Host "正在打包 exe（约 5–8 分钟）..."
    .\.venv\Scripts\pyinstaller.exe build.spec --noconfirm
}

# 6. 创建 Release 并上传附件
Write-Host "创建 Release 并上传 exe..."
gh release create $version `
    "dist\文档转Markdown.exe" `
    "dist\doc2md-cli.exe" `
    --title "文档转 Markdown $version" `
    --notes "Windows 桌面版 + 命令行版。详见 README 与 docs/使用说明.md。" `
    --latest

Write-Host ""
Write-Host "完成！" -ForegroundColor Green
Write-Host "仓库: https://github.com/$repo"
Write-Host "下载: https://github.com/$repo/releases/latest"
