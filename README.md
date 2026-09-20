# RFI01-Process raw tables from Asite

把 Asite 导出的 `Form Listing.xlsx` 整理成可用的**顾问分列**工作表。

## 做什么

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

## 文件

| 文件 | 说明 |
|---|---|
| `SKILL.md` | 技能定义（触发条件、流程、规则、踩坑） |
| `scripts/build_bs_listing.py` | 整理脚本 |
| `scripts/verify_output.py` | 校验脚本（含签名反查一致率） |

## 用法

```bash
pip install openpyxl

# 整理（默认读写 ~/Downloads/）
python scripts/build_bs_listing.py

# 指定路径
python scripts/build_bs_listing.py \
    --src "D:/x/Form Listing.xlsx" \
    --dst "D:/x/20260920-处理过Form Listing.xlsx"

# 校验
python scripts/verify_output.py "D:/x/20260920-处理过Form Listing.xlsx"
```

参数：`--src` `--dst` `--discipline BS`（可重复）`--keep-ar-str` `--date 20260920` `--sheet`

## 核心规则

- `Created_By` 是 `人名, 公司,人名, 公司, …`，公司代号只有 5 种：
  `Wong Tung`→WTP（建筑师）、`WSP`/`WSP (Asia) Ltd`→WSP（机电）、
  `MCS`→MSC（幕墙+结构）、`CSCE`→CSHK（总包，不落列）
- **按位配对**：顾问出现顺序 = 日期顺序 = 回复段落顺序
- 多段回复分隔符是 `" , "`；正文里也含该串时用带评分的动态规划挑切点
- 多日期分隔符必须是**纯逗号 `,`**，与源表 `Response_Date` 一致

## 准确性

用每段回复末尾署名（`Best Regards, 人名`）反查公司，与按位归属比对：
实测 **1936 个可判定段落 100% 一致**。
