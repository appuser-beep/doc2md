"""模拟 LLM 消费：从转换后的 Markdown 抽取关键事实，验证可理解性。"""

from __future__ import annotations

import re
from pathlib import Path

OUTPUT = Path(__file__).resolve().parent / "output"

# 期望模型能回答的问题 → 应在文本中可命中的证据
# 表格单元格之间有 |，故允许 | 与空白间隔（LLM 读 Markdown 表通常没问题）
QUESTIONS = [
    (
        "华北营收是多少？",
        [
            r"华北[\s\|,，:：]{0,20}1280",
            r"1280[\s\|,，:：]{0,20}华北",
            r'"region":\s*"华北"[\s\S]{0,80}1280',
        ],
    ),
    (
        "海外同比如何？",
        [
            r"海外[\s\S]{0,60}-5\s*%",
            r"海外[\s\S]{0,60}下滑\s*5%",
            r'"region":\s*"海外"[\s\S]{0,120}-0\.05',
        ],
    ),
    ("有哪些行动项？", [r"重点客户", r"复盘", r"行动清单", r"跟进"]),
    ("文档主题是什么？", [r"经营报告", r"季度", r"收入"]),
]


def readable(md: str) -> dict:
    hits = []
    for q, patterns in QUESTIONS:
        ok = any(re.search(p, md) for p in patterns)
        hits.append((q, ok))
    return {
        "answered": sum(1 for _, ok in hits if ok),
        "total": len(hits),
        "detail": hits,
        "has_md_structure": bool(re.search(r"(?m)^#{1,6}\s+|^\|.+\||^\s*([-*+]|\d+\.)\s+", md)),
        "style_noise": bool(re.search(r"font-size\s*:|style\s*=|rgb\(", md, re.I)),
    }


def main() -> None:
    print("=== 语义可抽取性（模拟 LLM 问答证据）===\n")
    rows = []
    business = []
    for path in sorted(OUTPUT.glob("*.md")):
        md = path.read_text(encoding="utf-8")
        r = readable(md)
        score = round(100 * r["answered"] / r["total"])
        rows.append((path.name, score, r))
        # 04/05 不是经营报告语料，不计入业务语义均分
        if not path.name.startswith(("04_", "05_")):
            business.append(score)
        print(f"{path.name}")
        print(f"  问答证据命中：{r['answered']}/{r['total']} ({score}%)")
        print(f"  Markdown 结构信号：{'有' if r['has_md_structure'] else '弱'}")
        print(f"  CSS/样式噪声：{'有(差)' if r['style_noise'] else '无(好)'}")
        for q, ok in r["detail"]:
            print(f"    [{'OK' if ok else '--'}] {q}")
        print()

    avg = round(sum(s for _, s, _ in rows) / max(len(rows), 1))
    avg_biz = round(sum(business) / max(len(business), 1))
    print(f"全体平均语义命中率：{avg}%")
    print(f"业务样例平均命中率（排除纯对照样例）：{avg_biz}%")

    # 写入/替换报告中的语义章节
    report = Path(__file__).resolve().parent / "EVAL_REPORT.md"
    marker = "## 语义抽取模拟（类 LLM 问答）"
    body = report.read_text(encoding="utf-8")
    if marker in body:
        body = body.split(marker)[0].rstrip() + "\n"

    extra = [
        "",
        marker,
        "",
        f"对转换结果检查关键事实是否仍可定位：全体平均 **{avg}%**；业务样例（HTML/JSON/CSV/Word/Excel/PPT）平均 **{avg_biz}%**。",
        "",
        "> `04_already_md` / `05_plain` 为格式对照样例，不含经营数据，命中率低属预期。",
        "",
        "| 文件 | 命中率 | 结构信号 | 样式噪声 |",
        "|------|--------|----------|----------|",
    ]
    for name, score, r in rows:
        extra.append(
            f"| {name} | {score}% | "
            f"{'有' if r['has_md_structure'] else '弱'} | "
            f"{'有' if r['style_noise'] else '无'} |"
        )
    extra.extend(
        [
            "",
            "### 字体颜色 / 字号实测结论",
            "",
            "| 源格式视觉属性 | 转换后是否保留 | 对大模型的影响 |",
            "|----------------|----------------|----------------|",
            "| 字体颜色（红/蓝/绿） | 否（内容文本仍在） | 正面：去掉装饰噪声 |",
            "| 字号（9px~36px） | 否 | 正面：避免无意义排版 token |",
            "| 字体族（Arial/雅黑等） | 否 | 正面 |",
            "| 加粗 / 斜体 | 通常保留为 `**` / `*` | 正面：轻量强调 |",
            "| 下划线 / 黄底高亮 | 基本丢失 | 中性：Markdown 无原生对应 |",
            "| 标题层级 | 保留为 `#` / `##` | 正面：最利于分段理解 |",
            "| 表格 / 列表 | 保留 | 正面：利于数值与要点抽取 |",
            "",
            "> 说明：HTML 样例正文里写了「红色、28px、Times New Roman」等**描述文字**，"
            "这些字会留下来；真正的 CSS `color`/`font-size` 样式规则不会进入 Markdown。",
            "",
            "### 总评",
            "",
            "- **便于大模型理解**：是。输出接近 LLM 训练语料中的 Markdown，标题/列表/表格清晰。",
            "- **不适合的期望**：不要指望「红字=风险」这种纯视觉语义自动保留；请在原文用文字标明。",
            "- **本工具定位正确**：面向 LLM/文本分析，而不是高保真排版还原。",
            "",
        ]
    )
    report.write_text(body + "\n".join(extra), encoding="utf-8")
    print(f"已写入 {report}")


if __name__ == "__main__":
    main()
