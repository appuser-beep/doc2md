# ZIP 增强 · 回归报告

**版本**：1.5.4  
**结果**：`tests.test_zip_convert` **6/6 OK**（与 Excel 合并用例合计 12/12）

## 根因（用户 `Java_SE_HomeWork.zip`）

文件扩展名为 `.zip`，但文件头为 `Rar!` —— **实际是 RAR**。  
官方 MarkItDown 只认 ZIP，旧版界面会卡在「正在转换」较久后才失败，体验像死机。

## 修复

1. 魔数嗅探：RAR/7z 伪 ZIP → 选文件/开始转换即明确报错  
2. 真 ZIP：本地遍历，转源码/文档，跳过 `.class`/`.jar`  
3. 进度条状态文案更细

## 用法建议

用 7-Zip / WinRAR 打开后「重新压缩为 ZIP」，或解压后直接转 `.java` / 文档。
