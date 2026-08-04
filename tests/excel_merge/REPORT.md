# Excel 宽合并折叠 · 回归报告

**版本**：1.5.3  
**结果**：`tests.test_excel_merge_collapse` **6/6 OK**

## 覆盖

| 用例 | 说明 |
| --- | --- |
| `test_wide_merge_row_*` | 横向刷屏折叠；两格碰巧相同不误压 |
| `test_dedupe_vertical_repeats` | 纵向合并填充后的连续重复行去重 |
| `test_synthetic_schedule` | 生成样例 `samples/schedule_wide_merge.xlsx` |
| `test_real_desktop_schedule_*` | 桌面原件「周一至周五档案整理.xlsx」 |
| `test_ex05_still_splits` | 并排表拆分不被折叠逻辑破坏 |

## 期望形态（排班表）

两列表，每时段一行，无横向/纵向刷屏：

```markdown
| 时间 | 周一到周五档案整理人员名单 |
| --- | --- |
| 周一上午 | 无 |
| 周一下午 | 王成，刘雨姗，郭家付，… |
| … | … |
```

输出快照：`tests/excel_merge/output/`。
