# 转换质量与 LLM 友好度测试报告

- 样例数：8
- 成功转换：8/8
- 结构保留均分：91/100
- LLM 友好均分：92/100

## 结论摘要

1. **字体颜色 / 字号**：MarkItDown（及 Markdown 本身）通常**不会保留**视觉样式；这反而更适合大模型——减少噪声 token，聚焦语义。
2. **便于 LLM 理解的部分**：标题、列表、表格、链接、关键数值与行动项大多能保留。
3. **需要注意**：纯装饰信息（红字强调、黄底高亮）会丢失；若业务依赖“颜色=风险等级”，应在源文档用文字写明（如“风险：红色/下滑”），不要只靠颜色。

## 逐项结果

### 01_styles_structure.html
- 输出：`tests/output/01_styles_structure.md`
- 结构分：100/100 · LLM 友好分：100/100
- 规模：571 字符 / 40 行
- 检查项：
  - ✓ `has_heading`
  - ✓ `has_list`
  - ✓ `has_table`
  - ✓ `has_key_semantics`
  - ✓ `no_css_noise`
  - ✓ `reasonable_length`
- 可读性备注：
  - 标题层级清晰，利于模型分段理解。
  - 表格以 Markdown 呈现，便于抽取数值。
  - 列表结构保留，适合作为行动项/要点输入。
  - 未保留字号/颜色等视觉样式（对 LLM 通常是优点）。

<details><summary>预览（前 1200 字符）</summary>

```markdown
# 一级标题：季度经营报告

## 二级标题：收入概况

### 三级标题：区域拆分

这段文字是红色、28px、Times New Roman —— 测试字体颜色与大小是否保留。

这段是蓝色小号 Arial 文字。

特大紫色加粗装饰文字

极小号灰色备注，对人眼几乎不可见，但应对 LLM 仍可读。

普通段落含 **加粗**、*斜体*、下划线、
黄底高亮 以及
[外部链接](https://example.com/report)。

* 无序列表 A：华北区增长 **12%**
* 无序列表 B：华南区持平
* 无序列表 C：海外新签 *3* 个客户

1. 优先跟进重点客户
2. 复盘渠道转化漏斗
3. 下周输出行动清单

| 区域 | 营收(万元) | 同比 |
| --- | --- | --- |
| 华北 | 1280 | +12% |
| 华南 | 960 | 0% |
| 海外 | 430 | -5% |

> 引用块：管理层要求关注海外下滑原因。

```
def calculate_growth(current, previous):
    return (current - previous) / previous
```

页脚装饰：© 2026 测试公司 · 字体颜色仅用于排版，不应影响语义。
```

</details>

### 02_metrics.json
- 输出：`tests/output/02_metrics.md`
- 结构分：100/100 · LLM 友好分：80/100
- 规模：394 字符 / 16 行
- 检查项：
  - ✓ `has_key_semantics`
  - ✓ `no_css_noise`
  - ✓ `reasonable_length`
- 可读性备注：
  - 缺少 Markdown 标题，模型需自行推断章节。
  - 未保留字号/颜色等视觉样式（对 LLM 通常是优点）。

<details><summary>预览（前 1200 字符）</summary>

```markdown
{
  "report_name": "季度经营报告",
  "quarter": "2026-Q2",
  "metrics": [
    {"region": "华北", "revenue_wan": 1280, "yoy": 0.12, "note": "重点增长区"},
    {"region": "华南", "revenue_wan": 960, "yoy": 0.0, "note": "持平"},
    {"region": "海外", "revenue_wan": 430, "yoy": -0.05, "note": "需复盘"}
  ],
  "actions": [
    "优先跟进重点客户",
    "复盘渠道转化漏斗",
    "输出下周行动清单"
  ],
  "owner": {"name": "张三", "role": "分析师"}
}

```

</details>

### 03_regions.csv
- 输出：`tests/output/03_regions.md`
- 结构分：100/100 · LLM 友好分：100/100
- 规模：178 字符 / 7 行
- 检查项：
  - ✓ `has_key_semantics`
  - ✓ `reasonable_length`
- 可读性备注：
  - 缺少 Markdown 标题，模型需自行推断章节。
  - 表格以 Markdown 呈现，便于抽取数值。
  - 未保留字号/颜色等视觉样式（对 LLM 通常是优点）。

<details><summary>预览（前 1200 字符）</summary>

```markdown
| 区域 | 营收(万元) | 同比 | 备注 |
| --- | --- | --- | --- |
| 华北 | 1280 | +12% | 重点增长区 |
| 华南 | 960 | 0% | 持平 |
| 海外 | 430 | -5% | 需复盘 |
| 华东 | 1100 | +8% | 稳定 |
| 西南 | 520 | +3% | 新开拓 |
```

</details>

### 04_already_md.md
- 输出：`tests/output/04_already_md.md`
- 结构分：75/100 · LLM 友好分：80/100
- 规模：231 字符 / 21 行
- 检查项：
  - ✓ `has_heading`
  - ✓ `has_list`
  - ✗ `has_key_semantics`
  - ✓ `reasonable_length`
- 可读性备注：
  - 标题层级清晰，利于模型分段理解。
  - 表格以 Markdown 呈现，便于抽取数值。
  - 列表结构保留，适合作为行动项/要点输入。
  - 未保留字号/颜色等视觉样式（对 LLM 通常是优点）。

<details><summary>预览（前 1200 字符）</summary>

```markdown
# 已有 Markdown 样例

这是一份 **已带结构** 的 Markdown，用于验证“Markdown → Markdown”是否保持可读。

## 要点

1. 标题层级应保留
2. 列表与表格应保留
3. 代码块应保留

| 指标 | 值 |
|------|----|
| 准确率 | 96% |
| 召回率 | 91% |

```python
print("hello llm")
```

> 引用：LLM 更关心语义结构，而非字号颜色。

```

</details>

### 05_plain.txt
- 输出：`tests/output/05_plain.md`
- 结构分：50/100 · LLM 友好分：80/100
- 规模：160 字符 / 17 行
- 检查项：
  - ✗ `has_key_semantics`
  - ✓ `reasonable_length`
- 可读性备注：
  - 缺少 Markdown 标题，模型需自行推断章节。
  - 列表结构保留，适合作为行动项/要点输入。
  - 未保留字号/颜色等视觉样式（对 LLM 通常是优点）。

<details><summary>预览（前 1200 字符）</summary>

```markdown
产品说明书（纯文本）

第一章 概述
本产品用于文档转换。字体颜色、字号在纯文本中不存在。

第二章 功能
1. 转换 PDF
2. 转换 Office
3. 转换网页

第三章 注意事项
- 大模型需要清晰标题与列表
- 表格数据应可读
- 避免无意义装饰符号堆砌

联系人：李四 / 电话：010-88886666

```

</details>

### 06_styled_report.docx
- 输出：`tests/output/06_styled_report.md`
- 结构分：100/100 · LLM 友好分：100/100
- 规模：228 字符 / 21 行
- 检查项：
  - ✓ `has_heading`
  - ✓ `has_list`
  - ✓ `has_table`
  - ✓ `has_key_semantics`
  - ✓ `no_css_noise`
  - ✓ `reasonable_length`
- 可读性备注：
  - 标题层级清晰，利于模型分段理解。
  - 表格以 Markdown 呈现，便于抽取数值。
  - 列表结构保留，适合作为行动项/要点输入。
  - 未保留字号/颜色等视觉样式（对 LLM 通常是优点）。

<details><summary>预览（前 1200 字符）</summary>

```markdown
# 季度经营报告（Word）

## 收入概况

**红色大号强调：华北增长 12%。**

蓝色小号说明：本段仅用于视觉排版测试。

普通正文：管理层要求复盘海外下滑原因。

* 重点客户跟进
* 渠道漏斗复盘

1. 输出行动清单

|  |  |  |
| --- | --- | --- |
| 区域 | 营收(万元) | 同比 |
| 华北 | 1280 | +12% |
| 华南 | 960 | 0% |
| 海外 | 430 | -5% |
```

</details>

### 07_styled_sheet.xlsx
- 输出：`tests/output/07_styled_sheet.md`
- 结构分：100/100 · LLM 友好分：100/100
- 规模：134 字符 / 6 行
- 检查项：
  - ✓ `has_key_semantics`
  - ✓ `reasonable_length`
- 可读性备注：
  - 标题层级清晰，利于模型分段理解。
  - 表格以 Markdown 呈现，便于抽取数值。
  - 未保留字号/颜色等视觉样式（对 LLM 通常是优点）。

<details><summary>预览（前 1200 字符）</summary>

```markdown
## 收入
| 区域 | 营收(万元) | 同比 | 备注 |
| --- | --- | --- | --- |
| 华北 | 1280 | +12% | 重点增长区 |
| 华南 | 960 | 0% | 持平 |
| 海外 | 430 | -5% | 需复盘 |
```

</details>

### 08_styled_deck.pptx
- 输出：`tests/output/08_styled_deck.md`
- 结构分：100/100 · LLM 友好分：100/100
- 规模：199 字符 / 13 行
- 检查项：
  - ✓ `has_key_semantics`
  - ✓ `no_css_noise`
  - ✓ `reasonable_length`
- 可读性备注：
  - 标题层级清晰，利于模型分段理解。
  - 表格以 Markdown 呈现，便于抽取数值。
  - 未保留字号/颜色等视觉样式（对 LLM 通常是优点）。

<details><summary>预览（前 1200 字符）</summary>

```markdown
<!-- Slide number: 1 -->
# 季度经营报告
华北增长 12%
海外下滑 5%（需复盘）
行动：跟进重点客户 / 复盘漏斗

<!-- Slide number: 2 -->
# 数据表
| 区域 | 营收 | 同比 |
| --- | --- | --- |
| 华北 | 1280 | +12% |
| 华南 | 960 | 0% |
| 海外 | 430 | -5% |
```

</details>

## 对工具使用的建议

| 场景 | 建议 |
|------|------|
| 给大模型做摘要/问答 | 直接用本工具输出的 Markdown，效果通常足够好 |
| 依赖颜色表达风险 | 在原文补充文字标签，勿只靠红/绿字 |
| 复杂扫描件 PDF | 可能缺版式；可后续接 OCR/云服务 |
| Token 控制 | 先转 Markdown 再切片，比直接塞 PDF 二进制更高效 |

## 机器可读结果

```json
[
  {
    "file": "01_styles_structure.html",
    "ok": true,
    "structure_score": 100,
    "llm_score": 100,
    "chars": 571,
    "error": null
  },
  {
    "file": "02_metrics.json",
    "ok": true,
    "structure_score": 100,
    "llm_score": 80,
    "chars": 394,
    "error": null
  },
  {
    "file": "03_regions.csv",
    "ok": true,
    "structure_score": 100,
    "llm_score": 100,
    "chars": 178,
    "error": null
  },
  {
    "file": "04_already_md.md",
    "ok": true,
    "structure_score": 75,
    "llm_score": 80,
    "chars": 231,
    "error": null
  },
  {
    "file": "05_plain.txt",
    "ok": true,
    "structure_score": 50,
    "llm_score": 80,
    "chars": 160,
    "error": null
  },
  {
    "file": "06_styled_report.docx",
    "ok": true,
    "structure_score": 100,
    "llm_score": 100,
    "chars": 228,
    "error": null
  },
  {
    "file": "07_styled_sheet.xlsx",
    "ok": true,
    "structure_score": 100,
    "llm_score": 100,
    "chars": 134,
    "error": null
  },
  {
    "file": "08_styled_deck.pptx",
    "ok": true,
    "structure_score": 100,
    "llm_score": 100,
    "chars": 199,
    "error": null
  }
]
```

## 语义抽取模拟（类 LLM 问答）

对转换结果检查关键事实是否仍可定位：全体平均 **72%**；业务样例（HTML/JSON/CSV/Word/Excel/PPT）平均 **96%**。

> `04_already_md` / `05_plain` 为格式对照样例，不含经营数据，命中率低属预期。

| 文件 | 命中率 | 结构信号 | 样式噪声 |
|------|--------|----------|----------|
| 01_styles_structure.md | 100% | 有 | 无 |
| 02_metrics.md | 100% | 弱 | 无 |
| 03_regions.md | 75% | 有 | 无 |
| 04_already_md.md | 0% | 有 | 无 |
| 05_plain.md | 0% | 有 | 无 |
| 06_styled_report.md | 100% | 有 | 无 |
| 07_styled_sheet.md | 100% | 有 | 无 |
| 08_styled_deck.md | 100% | 有 | 无 |

### 字体颜色 / 字号实测结论

| 源格式视觉属性 | 转换后是否保留 | 对大模型的影响 |
|----------------|----------------|----------------|
| 字体颜色（红/蓝/绿） | 否（内容文本仍在） | 正面：去掉装饰噪声 |
| 字号（9px~36px） | 否 | 正面：避免无意义排版 token |
| 字体族（Arial/雅黑等） | 否 | 正面 |
| 加粗 / 斜体 | 通常保留为 `**` / `*` | 正面：轻量强调 |
| 下划线 / 黄底高亮 | 基本丢失 | 中性：Markdown 无原生对应 |
| 标题层级 | 保留为 `#` / `##` | 正面：最利于分段理解 |
| 表格 / 列表 | 保留 | 正面：利于数值与要点抽取 |

> 说明：HTML 样例正文里写了「红色、28px、Times New Roman」等**描述文字**，这些字会留下来；真正的 CSS `color`/`font-size` 样式规则不会进入 Markdown。

### 总评

- **便于大模型理解**：是。输出接近 LLM 训练语料中的 Markdown，标题/列表/表格清晰。
- **不适合的期望**：不要指望「红字=风险」这种纯视觉语义自动保留；请在原文用文字标明。
- **本工具定位正确**：面向 LLM/文本分析，而不是高保真排版还原。
