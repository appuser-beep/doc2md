# 下载 GitHub CLI 到 tools/gh（无需管理员权限）
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$ver = "2.63.2"
$zip = Join-Path $root "tools\gh.zip"
$dest = Join-Path $root "tools\gh"
New-Item -ItemType Directory -Force -Path (Join-Path $root "tools") | Out-Null

Write-Host "下载 GitHub CLI v$ver ..."
Invoke-WebRequest -Uri "https://github.com/cli/cli/releases/download/v$ver/gh_${ver}_windows_amd64.zip" -OutFile $zip
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
Expand-Archive -Path $zip -DestinationPath $dest -Force
Remove-Item $zip -Force

$bin = Get-ChildItem $dest -Recurse -Filter gh.exe | Select-Object -First 1
Write-Host "已安装: $($bin.FullName)"
Write-Host "下一步: $($bin.FullName) auth login"
