# 一键发布到 GitHub（源码 + Release 附件）
# 用法：在仓库根目录执行 .\scripts\publish.ps1
#
# 首次使用前请先登录 GitHub（二选一）：
#   1) gh auth login
#   2) $env:GITHUB_TOKEN = "Personal Access Token"  （需 repo 权限）

$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)

$version = "v1.7.2"
$repo = "wluser3362203440/doc2md"

Write-Host "=== 文档转 Markdown 发布脚本 ===" -ForegroundColor Cyan
Write-Host "版本: $version"
Write-Host ""

function Find-Gh {
    $cmd = Get-Command gh -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $local = Join-Path $PSScriptRoot "..\tools\gh\bin\gh.exe"
    if (Test-Path $local) { return (Resolve-Path $local).Path }
    return $null
}

$ghPath = Find-Gh
if (-not $ghPath) {
    Write-Host "未检测到 GitHub CLI (gh)。" -ForegroundColor Yellow
    Write-Host "请运行: .\scripts\install-gh.ps1"
    Write-Host "然后:   .\tools\gh\bin\gh.exe auth login"
    exit 1
}

Write-Host "使用 gh: $ghPath"

if ($env:GITHUB_TOKEN) {
    $env:GITHUB_TOKEN | & $ghPath auth login --with-token 2>$null
}

& $ghPath auth status 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "未登录 GitHub。" -ForegroundColor Yellow
    Write-Host "请先执行: .\tools\gh\bin\gh.exe auth login"
    Write-Host "完成后再运行: .\scripts\publish.ps1"
    Write-Host "详见: scripts\发布说明.md"
    exit 1
}

& $ghPath repo view $repo 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "创建仓库 https://github.com/$repo ..."
    & $ghPath repo create $repo --public --source=. --remote=origin `
        --description "文档转 Markdown — Windows 工具，PDF/Word/Excel/PPT 等转 Markdown"
}

Write-Host "推送源码..."
git push -u origin main
if ($LASTEXITCODE -ne 0) {
    Write-Host "git push 失败。请确认已登录且仓库权限正确。" -ForegroundColor Red
    exit 1
}

git tag -a $version -m "Release $version" -f 2>$null
git push origin $version --force

if (-not (Test-Path "dist\doc2md-cli.exe")) {
    Write-Host "正在打包 exe（约 5–8 分钟）..."
    if (-not (Test-Path ".venv\Scripts\pyinstaller.exe")) {
        .\.venv\Scripts\pip.exe install pyinstaller
    }
    .\.venv\Scripts\pyinstaller.exe build.spec --noconfirm
}

if (-not (Test-Path "dist\doc2md-cli.exe")) {
    Write-Host "未找到 dist\doc2md-cli.exe，无法上传 Release 附件。" -ForegroundColor Red
    exit 1
}

Write-Host "创建 Release 并上传 exe..."
$notes = @"
文档转 Markdown $version

## 下载（Windows 免安装）

| 文件 | 说明 |
|------|------|
| 文档转Markdown.exe | 图形界面 |
| doc2md-cli.exe | 命令行（管道、批量） |

详细用法见仓库 docs/使用说明.md
"@

& $ghPath release view $version 2>$null
if ($LASTEXITCODE -eq 0) {
    & $ghPath release upload $version "dist\文档转Markdown.exe" "dist\doc2md-cli.exe" --clobber
    & $ghPath release edit $version --notes $notes
} else {
    & $ghPath release create $version `
        "dist\文档转Markdown.exe" `
        "dist\doc2md-cli.exe" `
        --title "文档转 Markdown $version" `
        --notes $notes `
        --latest
}

Write-Host ""
Write-Host "完成！" -ForegroundColor Green
Write-Host "仓库: https://github.com/$repo"
Write-Host "下载: https://github.com/$repo/releases/latest"
