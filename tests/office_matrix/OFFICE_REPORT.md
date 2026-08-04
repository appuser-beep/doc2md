# Office 穷举测试报告

- 时间：2026-08-03T17:43:18+08:00
- 合计：29 · 通过：29 · 失败：0

## WORD（10/10）

| 用例 | 结果 | 命中 | 表 | 图痕迹 | 说明 |
|------|------|------|----|--------|------|
| `W01_text_only` | PASS | 3/3 | N | N |  |
| `W02_text_table` | PASS | 4/4 | Y | N |  |
| `W03_text_image` | PASS | 3/3 | N | Y |  |
| `W04_text_image_table` | PASS | 4/4 | Y | Y |  |
| `W05_multi_table` | PASS | 4/4 | Y | N |  |
| `W06_multi_image` | PASS | 4/4 | N | Y |  |
| `W07_merged_table` | PASS | 4/4 | Y | N |  |
| `W08_header_footer` | PASS | 2/2 | N | N |  |
| `W09_hyperlink` | PASS | 2/2 | N | N |  |
| `W10_long_headings` | PASS | 4/4 | N | N |  |

## EXCEL（10/10）

| 用例 | 结果 | 命中 | 表 | 图痕迹 | 说明 |
|------|------|------|----|--------|------|
| `E01_text_only` | PASS | 2/2 | N | N |  |
| `E02_text_table` | PASS | 4/4 | Y | N |  |
| `E03_text_image` | PASS | 3/3 | N | Y |  |
| `E04_text_image_table` | PASS | 3/3 | Y | Y |  |
| `E05_multi_table_side` | PASS | 4/4 | Y | N |  |
| `E06_multi_sheet` | PASS | 3/3 | Y | N |  |
| `E07_merged` | PASS | 3/3 | Y | N |  |
| `E08_number_formats` | PASS | 1/1 | Y | N |  |
| `E09_formula` | PASS | 1/1 | Y | N |  |
| `E10_multi_image` | PASS | 2/2 | N | Y |  |

## PDF（9/9）

| 用例 | 结果 | 命中 | 表 | 图痕迹 | 说明 |
|------|------|------|----|--------|------|
| `P01_text_only` | PASS | 2/2 | N | N |  |
| `P02_text_table` | PASS | 3/3 | Y | N |  |
| `P03_text_image` | PASS | 3/3 | N | N | 未检测到图片痕迹（可能只保留图注文字） |
| `P04_text_image_table` | PASS | 4/4 | Y | N | 未检测到图片痕迹（可能只保留图注文字） |
| `P05_multi_table` | PASS | 3/3 | Y | N |  |
| `P06_multi_image` | PASS | 3/3 | N | N | 未检测到图片痕迹（可能只保留图注文字） |
| `P07_multipage` | PASS | 3/3 | N | N |  |
| `P08_scanned_like` | PASS | 0/0 | N | N | 扫描件无文字层（本地预期） |
| `P09_two_column` | PASS | 3/3 | N | N |  |
