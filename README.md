# 文档转 Markdown

**当前版本：v1.7.2 · Windows 桌面工具**

将 PDF、Word、Excel、PPT、网页、Jupyter 笔记本、ZIP 压缩包等常见格式转换为 Markdown，便于阅读、检索与后续处理。

- **项目主页：** https://github.com/wluser3362203440/doc2md  
- **详细使用说明：** [docs/使用说明.md](docs/使用说明.md)

---

## 下载安装（推荐）

无需安装 Python，下载 exe 即可使用。

👉 **最新版下载（Releases）：** https://github.com/wluser3362203440/doc2md/releases/latest

| 文件 | 大小约 | 用途 |
|------|--------|------|
| **文档转Markdown.exe** | 350 MB | 图形界面，双击打开，适合日常办公 |
| **doc2md-cli.exe** | 350 MB | 命令行，适合批量转换、脚本、管道 |

📄 图文说明见：[docs/使用说明.md](docs/使用说明.md)

**使用步骤：**

1. 打开上方 Releases 页面，下载对应 exe（两个文件能力相同，按需选择）。
2. 将 exe 放到任意文件夹（路径尽量不要含特殊字符）。
3. 双击 **文档转Markdown.exe** 即可开始转换。

> 首次启动可能稍慢（解压内置运行环境），属正常现象。Windows 若提示「未知发布者」，请选择「仍要运行」（软件未购买代码签名证书）。

---

## 图形界面快速上手

1. 点击 **选择文件**，或在下方的 URL 框输入网页地址  
2. 点击 **开始转换**，右侧预览 Markdown  
3. 点击 **另存为…** 保存 `.md` 文件  

**菜单说明：**

| 按钮 | 作用 |
|------|------|
| 格式说明 | 查看完整支持格式与限制 |
| 大模型设置 | 图片描述、内嵌图 OCR（需 API Key） |
| Azure 设置 | 扫描 PDF、视频、EML 等云端识别（需 Azure 账号） |
| 高级设置 | 样式映射、ExifTool、窄接口、保留内嵌图、自定义插件 |

配置文件保存在：`C:\Users\你的用户名\.doc2md\`

---

## 支持格式

| 类别 | 格式 |
|------|------|
| Office | Word (.docx)、Excel (.xlsx/.xls)、PPT (.pptx)、Outlook (.msg) |
| 文档 | PDF、EPUB、Jupyter (.ipynb) |
| 数据 | HTML、CSV、JSON、XML、RSS、Atom、TXT、Markdown |
| 媒体 | JPG/PNG、WAV/MP3/M4A、MP4 音轨 |
| 压缩包 | ZIP（遍历包内文档与源码） |
| 网络 | 网页、Wikipedia、Bing 搜索、YouTube 字幕 |

**云端扩展**（软件内配置后可用）：扫描 PDF、复杂版面、视频、EML、RTF 等。  
**不支持**：老格式 .doc / .ppt / .xlsb；RAR / 7z 直接转换。

---

## 命令行版（doc2md-cli.exe）

与图形版共用转换引擎与配置文件。

```powershell
# 单文件转换
doc2md-cli.exe 报告.pdf -o 报告.md

# 管道输入
type 报告.pdf | doc2md-cli.exe -x pdf -o 报告.md

# 窄接口（不使用 Excel / Notebook / ZIP 本地增强）
doc2md-cli.exe --local-only 文档.docx -o 文档.md

# 保留 PPT 内嵌 base64 图片
doc2md-cli.exe --keep-data-uris 幻灯片.pptx -o 幻灯片.md

# 查看已安装插件
doc2md-cli.exe --list-plugins
```

---

## 从源码运行（开发者）

### 环境

- Windows 10 / 11  
- Python 3.10 或以上  

### 安装

```powershell
git clone https://github.com/wluser3362203440/doc2md.git
cd doc2md
python -m venv .venv
.\.venv\Scripts\pip.exe install -r requirements.txt
```

### 启动

```powershell
# 图形界面
.\.venv\Scripts\python.exe main.py

# 命令行
.\.venv\Scripts\python.exe cli.py 样例.pdf -o 样例.md
```

### 自行打包 exe

```powershell
.\.venv\Scripts\pip.exe install pyinstaller
.\build.ps1
```

产物位于 `dist\文档转Markdown.exe` 与 `dist\doc2md-cli.exe`。

---

## Docker（命令行）

```powershell
docker build -t doc2md .
docker run --rm -v ${PWD}:/data doc2md 报告.pdf -o 报告.md
```

---

## 源码结构

```
doc2md/
├── main.py              启动入口（无参数→GUI，有参数→CLI）
├── app.py               图形界面
├── cli.py               命令行
├── converter.py         转换核心
├── excel_convert.py     Excel 增强（合并单元格、宽表）
├── ipynb_convert.py     Notebook 增强（保留输出）
├── zip_convert.py       ZIP 增强（魔数校验、跳过二进制）
├── cleanup.py           结果清理
├── llm_settings.py      大模型配置
├── azure_settings.py    Azure 配置
├── advanced_settings.py 高级选项
├── plugin_loader.py     插件加载
├── build.ps1            打包脚本
├── build.spec           PyInstaller 规格
├── Dockerfile
└── requirements.txt
```

---

## 常见问题

**转换后颜色、字号没了？**  
Markdown 保留结构与语义，不保留视觉样式；标题、列表、表格、链接会保留。

**扫描件 PDF 几乎没文字？**  
请在「大模型设置」或「Azure 设置」中启用 OCR / 文档智能服务。

**ZIP 提示实际是 RAR？**  
扩展名为 .zip 但内容可能是 RAR/7z，请重新压缩为真正的 ZIP。

**图形版和命令行版有什么区别？**  
转换能力相同；图形版便于手动操作，命令行版便于批量与自动化。

---

## 版本历史

| 版本 | 说明 |
|------|------|
| v1.7.2 | 保留内嵌图、插件列表、自定义插件、正式化格式说明 |
| v1.7.0 | 高级设置、CLI、Docker |
| v1.6.x | Azure / 大模型 GUI、Office 增强与回归测试 |

---

## 许可证

本项目采用 [MIT License](LICENSE) 开源。
