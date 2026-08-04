# L3/L4 条件测试报告

- 时间：2026-08-03T17:44:01+08:00
- 合计：4 · PASS：0 · FAIL：0 · SKIP：4

## 环境探测（不含密钥）

- L3 LLM 就绪：False · plugins：False · model：(未启用)
- L4 DocIntel 就绪：False · CU 就绪：False

| 用例 | 层级 | 状态 | 字符 | 耗时 | 说明 |
|------|------|------|------|------|------|
| `L3-01_llm_image_description` | L3 | **SKIP** | 0 | 0.0s | 未配置 LLM（需启用 + API Key） |
| `L3-02_markitdown_ocr_plugin` | L3 | **SKIP** | 0 | 0.0s | 未配置 LLM |
| `L4-01_docintel_scanned_pdf` | L4 | **SKIP** | 0 | 0.0s | 未配置 docintel_endpoint |
| `L4-02_cu_multimodal` | L4 | **SKIP** | 0 | 0.0s | 未配置 cu_endpoint |

## L3-01_llm_image_description
- 状态：**SKIP**
- 原因：未配置 LLM（需启用 + API Key）

## L3-02_markitdown_ocr_plugin
- 状态：**SKIP**
- 原因：未配置 LLM

## L4-01_docintel_scanned_pdf
- 状态：**SKIP**
- 原因：未配置 docintel_endpoint

## L4-02_cu_multimodal
- 状态：**SKIP**
- 原因：未配置 cu_endpoint

## 如何启用 L3/L4

**L3 大模型**：界面「大模型设置」→ 启用 + API Key + 可选 enable_plugins（OCR）。
或设置环境变量 `OPENAI_API_KEY`。

**L4 Azure**：界面「Azure 设置」→ docintel_endpoint / cu_endpoint + Key（或本机 `az login`）。
