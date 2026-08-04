# 发布到 GitHub（维护者）

本文说明如何将 **文档转 Markdown** 源码与 exe 发布到 GitHub，供他人下载使用。

## 一、首次准备

### 1. 创建 Personal Access Token（推荐）

1. 登录 GitHub → [Settings → Developer settings → Tokens](https://github.com/settings/tokens)
2. 选择 **Generate new token (classic)**
3. 勾选 **repo** 权限
4. 生成后复制 Token（只显示一次）

### 2. 安装 GitHub CLI

**方式 A — 官方安装包：** https://cli.github.com/

**方式 B — 项目脚本（免管理员）：**

```powershell
cd doc2md
.\scripts\install-gh.ps1
```

### 3. 登录

**方式 A — 浏览器登录：**

```powershell
gh auth login
```

**方式 B — Token 登录（适合脚本）：**

```powershell
$env:GITHUB_TOKEN = "你的Token"
gh auth login --with-token
# 粘贴 Token 后按 Ctrl+Z 再回车（PowerShell）
```

> GitHub **已不再支持**用账号密码直接 git push，请使用 Token 或 gh 登录。

## 二、一键发布

在仓库根目录执行：

```powershell
.\scripts\publish.ps1
```

脚本会自动：

1. 创建仓库 `wluser3362203440/doc2md`（若不存在）
2. 推送 main 分支源码
3. 打标签 `v1.7.2`
4. 若 `dist\` 无 exe 则自动打包
5. 创建 Release 并上传 **文档转Markdown.exe** 与 **doc2md-cli.exe**

## 三、发布后的地址

| 用途 | 链接 |
|------|------|
| 源码仓库 | https://github.com/wluser3362203440/doc2md |
| 最新版下载 | https://github.com/wluser3362203440/doc2md/releases/latest |
| 使用说明 | 仓库内 `docs/使用说明.md` |

## 四、仅推送源码（不上传 exe）

```powershell
git push -u origin main
git tag -a v1.7.2 -m "Release v1.7.2"
git push origin v1.7.2
```

推送标签后，GitHub Actions（`.github/workflows/release.yml`）可在云端自动打包并上传 Release（需仓库已启用 Actions）。

## 五、给用户看的文档

对外只维护以下两份（只介绍本工具）：

- 仓库首页：`README.md`
- 详细说明：`docs/使用说明.md`
