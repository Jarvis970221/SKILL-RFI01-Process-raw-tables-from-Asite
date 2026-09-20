---
name: rfi01-process-raw-tables-from-asite
description: RFI01-Process raw tables from Asite — process an Asite "Form Listing.xlsx" export into the tidy 顧問-split workbook: filter Discipline_Code to BS, drop AR/STR ref numbers, add the Asite search column, split the multi-consultant Response text and Response_Date into per-consultant WTP/WSP/MSC columns, natural-sort by Cont_Ref_No, and apply the yellow/pink column highlight scheme. Use when the user wants to 整理/处理/拆分 an Asite Form Listing export or rebuild "YYYYMMDD-处理过Form Listing.xlsx".
agent_created: true
---

# RFI01-Process raw tables from Asite

把 Asite 导出的 `Form Listing.xlsx` 整理成可用的**顾问分列**工作表。

## When to use

- 用户给了一个 Asite 导出的 `Form Listing.xlsx`（或从 Asite 重新导出），要求整理/处理/拆分
- 用户提到 `Discipline_Code`、`Created_By`、`Response_Date`、`Cont_Ref_No`、`Asite Reply` 这类字段
- 用户要求重建 `YYYYMMDD-处理过Form Listing.xlsx`

## Quick start

```bash
PY="<python>"                    # 需 openpyxl：pip install openpyxl
SK="<skills>/rfi01-process-raw-tables-from-asite/scripts"

# 整理（默认读 <下载目录>/Form Listing.xlsx，
#      输出 <下载目录>/<YYYYMMDD>-处理过Form Listing.xlsx）
"$PY" "$SK/build_bs_listing.py"

# 指定输入输出
"$PY" "$SK/build_bs_listing.py" --src "D:/x/Form Listing.xlsx" --dst "D:/x/20260920-处理过Form Listing.xlsx"

# 校验输出
"$PY" "$SK/verify_output.py" "D:/x/20260920-处理过Form Listing.xlsx"
```

常用参数：`--src` `--dst` `--discipline BS`（可重复）、`--keep-ar-str`（保留 AR/STR 单号）、
`--date 20260920`（输出文件名日期）、`--sheet`（工作表名，默认 `Form Listing`）。

依赖：`openpyxl`（`pip install openpyxl`）。

## 源表结构（重要）

- 单工作表 `Form Listing`
- **第 7 行才是表头**（前 6 行是导出信息），数据从**第 8 行**开始
- 原始 **14 列**：

```
最后更新(CST) / 创建日期(CST) / 旗子 / 表单标题 / 状态 / 关联 /
Cont_Ref_No / Building-Location / Floor / Discipline_Code /
Response / Created_By / Response_Date / Response_Flag
```

## 处理步骤

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

### ①②③ 筛选

- `Discipline_Code` 只保留 `BS`
- 再剔除 `Cont_Ref_No` 以 `NAH-A/CS/RFI/AR/` / `NAH-A/CS/RFI/STR/` 开头的行
  —— 这些行虽被标成 BS，但单号学科段是建筑(AR)/结构(STR)
- **保留** `BS(AC)` / `BS(FS)` / `BS(ST)` / `CSF(SD)/BS(O)` / `NAF-A/CS/RFI/BS/` 等 BS 变体

### ③ 解析 Created_By

结构是 `人名, 公司,人名, 公司, …`（逗号分隔，**人名在前、公司名在后**）。公司代号只有 5 种：

| 源写法 | 代号 | 角色 |
|---|---|---|
| `Wong Tung` | **WTP** | 建筑师（统筹方） |
| `WSP (Asia) Ltd` / `WSP` | **WSP** | MEP 机电顾问 |
| `MCS` | **MSC** | 幕墙 + 结构 |
| `CSCE` | CSHK | 总包（**不落列**） |

45 个人名→公司映射无歧义，脚本从数据里自动反推（`Leo Lee/Jacky Hui/Kit Lam → WSP`、
`Zoe Li/Bonnie Keung/Rex Choi → WTP`、`KC Hon/Kenneth LAW → MCS`）。

### ④⑤ 按位配对（核心）

`Created_By` 中顾问出现的**顺序（左→右）** = `Response_Date` 日期顺序 = `Response` 段落顺序。

- 同一顾问出现两次 → 两段回复用 `\n\n----------\n\n` 连接；两个日期用**纯逗号**连接
- 该顾问未出现 → 留空
- CSHK 的段落丢弃

**日期分隔符必须是纯逗号 `,`**，与源表 `Response_Date` 保持一致（`2023-09-07,2023-09-09`）。
不要写成 ` / ` 或 `, `。

### ⑤ Response 切分

多段回复用 **`" , "`（空格+逗号+空格）** 拼接。

1. 先 `split(' , ')`，段数若等于 `Created_By` 条目数 → 采用
2. 不等 → 用**带评分的动态规划**挑 n-1 个切点。评分依据：
   - 切点后以 `Dear` / 人名 / `WTPL|WSP|MCS|CSHK` / `Close` 开头 → 加分
   - 切点前以句号或人名结尾 → 加分
   - 切点后是小写字母、`(`、`-`、罗马数字、阿拉伯数字开头 → 减分

**为什么需要 DP**：正文里本身含 `" , "`（如 `…Block G/F , for further coordination`），
直接 split 会多切。实测 3338 行里 3327 行可直接切，10 行需 DP。

### ⑦ Asite search

插在 `旗子` 与 `表单标题` 之间（D 列），取**表单标题开头连续数字**（多为 5 位），
写成字符串 + `number_format='@'` 保留前导 0（`00873`、`00108`）。标题开头无数字则留空。

### ⑧ 填色与边框

| 颜色 | 色值 | 列（1-based） |
|---|---|---|
| 淡黄 `#FFF2CC` | `4, 13, 14, 15, 17, 18, 19` | Asite search、Asite Reply ×3、Asite Reply Date ×3 |
| 淡粉 `#FCE4EC` | `5, 6, 8, 9, 10, 20` | 表单标题、状态、Cont_Ref_No、Building-Location、Floor、Response_Date |

边框 `thin / #808080`，表头行与全部数据行一起填色。冻结窗格 `A8`。

## 最终 21 列

```
 1 最后更新 (CST)          11 Discipline_Code
 2 创建日期 (CST)          12 Response
 3 旗子                    13 Asite Reply (WTP)      淡黄
 4 Asite search   淡黄 @   14 Asite Reply (WSP)      淡黄
 5 表单标题        淡粉     15 Asite Reply (MSC)      淡黄
 6 状态           淡粉     16 Created_By
 7 关联                   17 Asite Reply Date (WTP)  淡黄
 8 Cont_Ref_No    淡粉     18 Asite Reply Date (WSP)  淡黄
 9 Building-Location 淡粉  19 Asite Reply Date (MSC)  淡黄
10 Floor          淡粉     20 Response_Date          淡粉
                          21 Response_Flag
```

## 验证（必做）

跑 `scripts/verify_output.py`，重点看 **签名反查**：把每段回复末尾的署名
（`Best Regards / Regards / Thanks, 人名`）反查人名→公司，与落位列比对。

实测 **1936 个可判定段落 100% 一致**。若一致率明显低于 100%，说明切分或按位配对出错，
不要交付，回头查 DELIM / COMP / 列索引。

## 踩过的坑

| 现象 | 原因 / 解法 |
|---|---|
| `save` 抛 `PermissionError` | 输出文件被 Excel/WPS 打开，目录会生成 `~$<文件名>` 锁文件。先确认锁消失再重跑 |
| 排序结果不对 | 插入 `Asite search` 后 `Cont_Ref_No` 索引从 6 变 **7**，排序键别写错 |
| Response 切分段数偏多 | 正文自带 `" , "`，用 DP 兜底 |
| 签名比对出现"不一致" | 大小写问题（`Rex Choi` 常写成 `rex choi`），匹配前统一 lowercase |
| 表头行错位 | 源表前 6 行是导出信息，表头在**第 7 行** |

## 已知瑕疵（交付时如实说明）

1. 表头按用户口径写成 `MSC`，但 `Created_By` 里实际代号是 `MCS`（保持用户写法）
2. 少数行 Response 末段是孤立的 `refer to attached`（附件链接文字），按位落到了 WTP，
   语义上可能属于前一段
3. CSHK（总包）的回复/日期不落任何列，被丢弃
4. 个别行 `Cont_Ref_No` 被 Asite 误填成表单标题；排序时非 `NA[FH]` 开头的排到末尾
