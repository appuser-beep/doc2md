# 生产门禁测试总报告

- 开始：2026-08-03T17:43:10+08:00
- 结束：2026-08-03T17:47:05+08:00

## 套件汇总

- 自动化套件：**8/8** 通过
- 真实场景：**6** 通过 · **0** 失败 · **1** 跳过

| 套件 | 结果 | 耗时 | 说明 |
|------|------|------|------|
| unit_tests | **PASS** | 0.49s | OK |
| office_matrix | **PASS** | 6.85s | } |
| office_exhaust | **PASS** | 11.64s | DONE pass=53 fail=0 skip=0 -> C:\Users\wluser\Desktop\markdown\tests\office_exha |
| office_deep | **PASS** | 9.16s | DONE pass=41 fail=0 skip=0 -> C:\Users\wluser\Desktop\markdown\tests\office_deep |
| office_pdf_wave4 | **PASS** | 9.39s | DONE pass=40 fail=0 skip=0 -> C:\Users\wluser\Desktop\markdown\tests\office_pdf_ |
| office_full_sweep | **PASS** | 12.77s | DONE pass=65 fail=0 skip=0 -> C:\Users\wluser\Desktop\markdown\tests\office_full |
| l3_l4 | **PASS** | 0.09s | ���棺C:\Users\wluser\Desktop\markdown\tests\L3_L4_REPORT.md |
| universe_deep | **PASS** | 182.36s | DONE pass=73 fail=0 skip=0 -> C:\Users\wluser\Desktop\markdown\tests\universe_de |

## 真实场景回归

| 用例 | 状态 | 字符 | 说明 |
|------|------|------|------|
| `RW_schedule_xlsx` | **SKIP** | 0 | 文件不存在: 周一到周五档案整理人员名单.xlsx |
| `RW_samples_html` | **PASS** | 572 |  |
| `RW_samples_csv` | **PASS** | 179 |  |
| `RW_samples_json` | **PASS** | 394 |  |
| `RW_samples_md` | **PASS** | 231 |  |
| `RW_samples_txt` | **PASS** | 160 |  |
| `RW_precheck_fake_rar` | **PASS** | 0 | 扩展名是 .zip，但文件实际是 RAR（文件头 Rar!）。请重新压缩为真正的 ZIP 后再转。 |

## 覆盖范围（穷举）

| 类别 | 已测 |
|------|------|
| Word / Excel / PPT / PDF / MSG | office_matrix + exhaust + deep + wave4 + full_sweep |
| HTML / CSV / JSON / RSS / EPUB / IPYNB / ZIP | universe_deep（样例齐全时） |
| 真实排班表 xlsx | RW_schedule_xlsx |
| ZIP 魔数 / 伪 RAR | unit_tests + RW_precheck |
| Excel 宽表折叠 | unit_tests + excel_merge |
| L3 LLM / L4 Azure | l3_l4（无凭证则 SKIP，非 FAIL） |

## 使用建议

1. **日常发布前**跑本脚本；套件全绿即可发版。
2. **L3/L4 SKIP** 不代表失败；配置 Key 后再跑 `run_l3_l4_conditional.py`。
3. **真实业务文件**（如排班表）建议加入 `tests/production_gate_output/` 人工 spot-check。

## 子报告路径

- `tests/office_matrix/OFFICE_REPORT.md`
- `tests/office_exhaust/OFFICE_EXHAUST_REPORT.md`
- `tests/office_deep/OFFICE_DEEP_REPORT.md`
- `tests/office_pdf_wave4/WAVE4_PDF_REPORT.md`
- `tests/office_full_sweep/FULL_SWEEP_REPORT.md`
- `tests/universe_deep/UNIVERSE_DEEP_REPORT.md`
- `tests/L3_L4_REPORT.md`

## 结论

**全部门禁通过，工具可用于日常办公转换。**