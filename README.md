# RFI01-Process raw tables from Asite

把 Asite 导出的 `Form Listing.xlsx` 整理成可用的**顾问分列**工作表，并支持**周度对比更新标记**。

## 做什么

### 第 1 步：整理

```
源表 6137 行
 ├─① 筛 Discipline_Code = BS               → 3338
 ├─② 删 Cont_Ref_No 的 AR/STR 单号          → 3295
 ├─③ 解析 Created_By → 顾问序列
 ├─④ 日期按位配对（逗号分隔）
 ├─⑤ 回复按位分列（" , " 切分 + DP 兜底）
 ├─⑥ 按 Cont_Ref_No 自然排序
 ├─⑦ 新增 Asite search 列（文本格式）
 └─⑧ 淡黄 7 列 + 淡粉 6 列，细灰边框
→ 成果 3295 行 × 21 列
```

### 第 2 步：周度对比（可选）

与上一期成品逐行逐格比对，在最左侧插入 3 列标记，并把**有差异的格子标亮黄**：

```
 ├─ (RFI有更新)   有任一更新才填当日日期
 ├─ (RFI状态更新) 新增 / 新Close
 └─ (RFI回复更新) WTP新回复 / WSP新回复 / MSC新回复
→ 24 列；实测 20260922：59 行标记、405 个亮黄格
```

## 文件

| 文件 | 说明 |
|---|---|
| `SKILL.md` | 技能定义（触发条件、流程、规则、踩坑） |
| `scripts/build_bs_listing.py` | 第 1 步：整理脚本 |
| `scripts/verify_output.py` | 校验脚本（含签名反查一致率） |
| `scripts/02_add_update_columns.py` | 第 2 步：更新标记 + 亮黄差异高亮 |
| `references/20260920-FormListing处理流程总结.md` | 2026-09-20 实测流程、统计与已知瑕疵 |

## 用法

```bash
pip install openpyxl

# 第 1 步：整理（默认读写 ~/Downloads/）
python scripts/build_bs_listing.py
python scripts/build_bs_listing.py \
    --src "D:/x/Form Listing.xlsx" \
    --dst "D:/x/20260922-处理过Form Listing.xlsx" --date 20260922
python scripts/verify_output.py "D:/x/20260922-处理过Form Listing.xlsx"

# 第 2 步：与上一期成品对比，加 3 列标记 + 亮黄高亮
python scripts/02_add_update_columns.py \
    --cur "D:/x/20260922-处理过Form Listing.xlsx" \
    --prev "D:/x/20260920-处理过Form Listing.xlsx" \
    --dst "//10.149.8.8/机电/.../260922 vs 260920 RFI Record_RFI Information/20260922-处理过Form Listing.xlsx" \
    --date 2026-09-22
```

参数：`--src` `--dst` `--discipline BS`（可重复）`--keep-ar-str` `--date 20260922` `--sheet`
／ `--cur` `--prev` `--dst` `--date 2026-09-22`

## 核心规则

- `Created_By` 是 `人名, 公司,人名, 公司, …`，公司代号只有 5 种：
  `Wong Tung`→WTP（建筑师）、`WSP`/`WSP (Asia) Ltd`→WSP（机电）、
  `MCS`→MSC（幕墙+结构）、`CSCE`→CSHK（总包，不落列）
- **按位配对**：顾问出现顺序 = 日期顺序 = 回复段落顺序
- 多段回复分隔符是 `" , "`；正文里也含该串时用带评分的动态规划挑切点
- 多日期分隔符必须是**纯逗号 `,`**，与源表 `Response_Date` 一致
- 周度对比必须**成品 vs 成品**，且同单号多行要用**签名配对**（两期行序会互换）
- 第 2 步的坑：先在本期原始列位取值缓存，再 `insert_cols(1,3)`，否则全表误判

## 准确性

用每段回复末尾署名（`Best Regards, 人名`）反查公司，与按位归属比对：
实测 **1936 个可判定段落 100% 一致**（20260922 期 1979 段同样 100%）。
