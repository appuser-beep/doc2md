# Publish source and Release assets to GitHub
# Usage: .\scripts\publish.ps1  (from repo root)
#
# First-time: gh auth login  OR  $env:GITHUB_TOKEN = "PAT with repo scope"

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Set-Location (Split-Path $PSScriptRoot -Parent)

$version = "v1.7.9"
$repo = "appuser-beep/doc2md"
$repoUrl = "https://github.com/$repo"
$releaseUrl = "$repoUrl/releases/latest"

Write-Host "=== doc2md publish ===" -ForegroundColor Cyan
Write-Host "Version: $version"
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
    Write-Host "GitHub CLI (gh) not found." -ForegroundColor Yellow
    Write-Host "Run: .\scripts\install-gh.ps1"
    Write-Host "Then: .\tools\gh\bin\gh.exe auth login"
    exit 1
}

Write-Host "Using gh: $ghPath"

if ($env:GITHUB_TOKEN) {
    $env:GITHUB_TOKEN | & $ghPath auth login --with-token 2>$null
}

& $ghPath auth status 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Not logged in to GitHub." -ForegroundColor Yellow
    Write-Host "Run: .\tools\gh\bin\gh.exe auth login"
    exit 1
}

& $ghPath repo view $repo 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Creating repo $repoUrl ..."
    $desc = "doc2md - PDF/Word/Excel/PPT to Markdown for Windows"
    & $ghPath repo create $repo --public --source=. --remote=origin --description $desc
}

Write-Host "Pushing source..."
git push -u origin main
if ($LASTEXITCODE -ne 0) {
    Write-Host "git push failed. Check login and repo permissions." -ForegroundColor Red
    exit 1
}

git tag -a $version -m "Release $version" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Tag $version already exists or could not be created." -ForegroundColor Red
    Write-Host "Bump APP_VERSION / scripts/publish.ps1 version — do not force-move release tags." -ForegroundColor Yellow
    exit 1
}
git push origin $version
if ($LASTEXITCODE -ne 0) {
    Write-Host "Failed to push tag $version." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path "dist\doc2md-cli.exe")) {
    Write-Host "Building exe (about 5-8 min)..."
    if (-not (Test-Path ".venv\Scripts\pyinstaller.exe")) {
        .\.venv\Scripts\pip.exe install pyinstaller
    }
    .\.venv\Scripts\pyinstaller.exe build.spec --noconfirm
}

if (-not (Test-Path "dist\doc2md-cli.exe")) {
    Write-Host "dist\doc2md-cli.exe not found; cannot upload Release assets." -ForegroundColor Red
    exit 1
}

Write-Host "Creating Release and uploading exe..."
$notes = @"
doc2md $version

## Downloads (Windows, no install)

| File | Description |
|------|-------------|
| doc2md-gui.exe (文档转Markdown.exe) | GUI |
| doc2md-cli.exe | CLI / pipes / batch |

See docs/使用说明.md in the repo.
"@

$guiExe = Get-ChildItem -Path "dist" -Filter "*Markdown.exe" | Select-Object -First 1
if (-not $guiExe) {
    Write-Host "GUI exe not found under dist\" -ForegroundColor Red
    exit 1
}

& $ghPath release view $version 2>$null
if ($LASTEXITCODE -eq 0) {
    & $ghPath release upload $version $guiExe.FullName "dist\doc2md-cli.exe" --clobber
    & $ghPath release edit $version --notes $notes
} else {
    & $ghPath release create $version `
        $guiExe.FullName `
        "dist\doc2md-cli.exe" `
        --title "doc2md $version" `
        --notes $notes `
        --latest
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host ("Repo:    {0}" -f $repoUrl)
Write-Host ("Release: {0}" -f $releaseUrl)
