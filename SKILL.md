---
name: rfi01-process-raw-tables-from-asite
description: 'RFI01-Process raw tables from Asite — process an Asite "Form Listing.xlsx" export into the tidy 顧問-split workbook: filter Discipline_Code to BS, drop AR/STR ref numbers, add the Asite search column, split the multi-consultant Response text and Response_Date into per-consultant WTP/WSP/MSC columns, natural-sort by Cont_Ref_No, and apply the yellow/pink column highlight scheme. Weekly step 2 adds three update-marker columns ((RFI有更新)/(RFI状态更新)/(RFI回复更新)) versus last period and highlights every changed cell in bright yellow. Use when the user wants to 整理/处理/拆分 an Asite Form Listing export, rebuild "YYYYMMDD-处理过Form Listing.xlsx", or 做本期更新标记/对比上一期.'
agent_created: true
---

# RFI01-Process raw tables from Asite

把 Asite 导出的 `Form Listing.xlsx` 整理成可用的**顾问分列**工作表。

## When to use

- 用户给了一个 Asite 导出的 `Form Listing.xlsx`（或从 Asite 重新导出），要求整理/处理/拆分
- 用户提到 `Discipline_Code`、`Created_By`、`Response_Date`、`Cont_Ref_No`、`Asite Reply` 这类字段
- 用户要求重建 `YYYYMMDD-处理过Form Listing.xlsx`
- 用户要做**本期更新标记 / 与上一期对比 / 差异标亮黄 / 周度对比**

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

# 第 2 步（周度对比）：给本期成品加 3 列更新标记，并把与上期有差异的格子标亮黄；
#   成品存回本期原表所在目录
"$PY" "$SK/02_add_update_columns.py" \
    --cur "D:/x/20260922-处理过Form Listing.xlsx" \
    --prev "D:/x/20260920-处理过Form Listing.xlsx" \
    --dst "//10.149.8.8/机电/01 啟德醫院、中醫院/1 啟德醫院/5 RFI/260922 vs 260920 RFI Record_RFI Information/20260922-处理过Form Listing.xlsx" \
    --date 2026-09-22
```

### 第 2 步：周度更新标记（02_add_update_columns.py）

对本期 21 列成品 `insert_cols(1,3)`，新增三列（表头行第 7 行，冻结改 `D8`）：

| 列 | 规则 |
|---|---|
| `(RFI有更新)` | 该行命中任一更新（新增 / 新Close / 顾问新回复）才填 `--date`（默认当天，YYYY-MM-DD 文本）；无变化留空 |
| `(RFI状态更新)` | 上期没有该 `Cont_Ref_No` → `新增`；本期 `Closed` 且上期非 Closed → `新Close`；两期都 Closed 但 `Response_Date` 变了 → 也填 `新Close`；其余不填 |
| `(RFI回复更新)` | `Asite Reply Date (WTP/WSP/MSC)` 任一列值变化 → `WTP新回复` / `WSP新回复` / `MSC新回复`，多个用全角 `；` 连接 |

**同单号多行必须用"签名配对"**（2026-09-22 踩坑）：同一 `Cont_Ref_No` 有多行
（Old/New flag），且**两期成品里行序会互换**。按位置 i-th 对 i-th 会误标
（实测 88 行误标 → 签名配对后收敛为 53 行真变化）。做法：先按
`(状态,WTP日期,WSP日期,MSC日期,Response_Date)` 签名把**完全相同的行**消耗配对，
剩余行再按序配对、与"空"配对（视为该 RFI 多出的行）。

成品 → 成品直接对比（本期 21 列成品 vs 上期 21 列成品），**不要**用网络盘的
20 列模板当对比基准——那种模板的 Response/状态可能停留在更早时点。

20260922 实测：新增 0、新Close 0（Closed 两期均 2172；53 行 Open→Responded）、
WSP 新回复 53 行，共标记 53+6=59 行。

#### 亮黄差异高亮（2026-09-22 新增）

除了三列标记，脚本还会把**与上一期值不同的每一个单元格**填亮黄色，便于人工逐格核对：

- 色值 **`FFFFFF00`（亮黄）**，与已有的淡黄 `FFF2CC`、淡粉 `FCE4EC` 区分
- 对比范围：**21 个原数据列全部逐格比对**（不比对新增的 3 个标记列）
- 新增行（上期没有该单号）→ 整行 21 格标亮黄
- 上期成品的列位**按表头名定位**（`header_map()` 跳过 3 个标记列），
  因此上期成品是 21 列还是已插过标记列的 24 列都能正确对齐

20260922 实测 59 行、**405 个亮黄格**，按列分布：

| 列 | 格数 | | 列 | 格数 |
|---|---|---|---|---|
| 最后更新 (CST) | 59 | | 状态 | 53 |
| Response | 57 | | Created_By | 53 |
| Asite Reply (WSP) | 55 | | Asite Reply Date (WSP) | 53 |
| Response_Date | 53 | | 创建日期 6 / Asite search 4 / 表单标题 4 / Floor 4 / WTP回复 2 / Building-Location 2 | 零散 |

「最后更新 (CST)」59 格 > 回复行 53 行属正常：Asite 侧被触碰过的行时间戳会变，
即使语义上没有新回复。

**下游衔接（第 3 步）**：亮黄格是下一环的输入——技能 **`rfi02-sync-listing-to-schedule-base`**
会读取本期成品里的亮黄格，把「值与飞书 Base 不同」的格子同步进 `BS RFI SCHEDULE` 多维表
（`tblpOUNvLVXeUFTT`）。因此第 2 步必须用**标准亮黄 FFFFFF00**，不要改色值。


常用参数：`--src` `--dst` `--discipline BS`（可重复）、`--keep-ar-str`（保留 AR/STR 单号）、
`--date 20260920`（输出文件名日期）、`--sheet`（工作表名，默认 `Form Listing`）。

依赖：`openpyxl`（`pip install openpyxl`）。

**本机解释器选择（2026-09-21 实测）**：托管版本 `binaries/python/versions/3.13.12/python.exe`
**没有** openpyxl，直接用会 `ModuleNotFoundError`。请改用已装好依赖的托管 venv：

```
PY="C:/Users/nomo/.workbuddy/binaries/python/envs/default/Scripts/python.exe"   # openpyxl 3.1.5 ✅
```

若该 venv 也不可用，再退回系统 Python
`C:/Users/nomo/AppData/Local/Programs/Python/Python313/python.exe`（自带 openpyxl 3.1.5）。
报错信息里若出现 `No module named 'openpyxl'`，先换解释器，不要去改脚本。

> 另注：bash 的 PATH 偶发失效（`ls`/`grep` 报 `command not found`）时，
> 用**绝对路径**调用 python 即可，不要依赖 `cd`/管道。

## 周度作业流程（网络盘标准路径）

原表与成品都在网络盘，**每周一个对比目录**：

```
//10.149.8.8/机电/01 啟德醫院、中醫院/1 啟德醫院/5 RFI/
├── 260920 vs 260911 RFI Record_RFI Information/
│     ├── Form Listing.xlsx                      ← 该期原始导出（14 列）
│     └── 20260920-处理过Form Listing.xlsx        ← 成品（21 列；本期作为对比基准）
├── 260922 vs 260920 RFI Record_RFI Information/
│     ├── Form Listing.xlsx                      ← 本期原表
│     └── 20260922-处理过Form Listing.xlsx        ← 本期成品（24 列，含标记列+亮黄）
└── …
```

标准动作：

1. 让用户/自己确认**本期目录**与**上一期成品**路径（对 UNC 路径用 `//10.149.8.8/...` 正斜杠）
2. `build_bs_listing.py --src <本期目录>/Form Listing.xlsx --dst <本期目录>/<YYYYMMDD>-处理过Form Listing.xlsx --date YYYYMMDD`
3. `verify_output.py <本期成品>`
4. `02_add_update_columns.py --cur <本期成品> --prev <上期目录>/<上期成品> --dst <本期成品>`（覆盖）
5. 交付时说明：标记行数、亮黄格数、以及是否真有新增/新 Close

**成品默认放回本期原表目录**（用户明确要求），不要只留在 Downloads。

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
| 亮黄 `#FFFF00` | 第 2 步按差异**逐格**加 | 与上期值不同的任意数据格（见「亮黄差异高亮」） |

边框 `thin / #808080`，表头行与全部数据行一起填色。冻结窗格 `A8`（第 2 步插列后改 `D8`）。

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

第 2 步后变成 **24 列**：最前面插入 `(RFI有更新)` / `(RFI状态更新)` / `(RFI回复更新)`，
其余 21 列整体右移 3 位（底色、边框、`@` 文本格式随值一起右移，列宽不自动平移需手动设）。

## 验证（必做）

跑 `scripts/verify_output.py`，重点看 **签名反查**：把每段回复末尾的署名
（`Best Regards / Regards / Thanks, 人名`）反查人名→公司，与落位列比对。

实测 **1936 个可判定段落 100% 一致**。若一致率明显低于 100%，说明切分或按位配对出错，
不要交付，回头查 DELIM / COMP / 列索引。

## 20260920 实测基线

以下数字来自 `20260920-FormListing处理流程总结.md`，用于判断一次重跑是否落在预期范围内：

- 源表 6137 行；筛选 `BS` 后 3338 行；剔除 AR/STR 单号后 3295 行
- 成果为 21 列；`Asite search` 有值 3294 行，1 行为空
- `Asite Reply` 非空：WTP 1681、WSP 2635、MSC 204；三项全空 276 行
- 同一顾问多段回复合并 505 行；含多日期的单元格 534 个
- 切分统计：3338 行中 3327 行可直接按 `" , "` 切分，10 行需要 DP 兜底
- 签名反查 1936 个可判定段落，预期一致率为 100%

数字因源表版本不同可能变化；应优先检查筛选条件、列索引、分隔符和校验脚本输出，不要为了追数字修改源数据。

### 20260921 / 20260922 实测

| 项 | 20260921 | 20260922 |
|---|---|---|
| 源表数据行 | 6137 | 6138（+1 行 STR） |
| 筛 `BS` 后 | 3338 | 3338 |
| 剔 AR/STR 后 | 3295 | 3295 |
| Asite search 有值 | 3294 | 3294 |
| Reply 非空 WTP / WSP / MSC | 1681 / 2635 / 204 | 1681 / **2688** / 204 |
| 三项全空 | 276 | **223** |
| 签名反查一致率 | 100% (1927) | 100% (1979) |
| 第 2 步标记 | — | 59 行 / 405 亮黄格（WSP 新回复 53） |

20260921 与 20260920 的内容完全等同（逐项吻合），说明这两次导出是同一版；
20260922 有 53 行 WSP 新回复（主要是 03116–03175 那批 `Request Architect's Formal Reply`
SOA 批量 RFI，状态 Open→Responded）。

> 注意：用「网络盘 20 列模板」当对比基准会得出**虚高**的变更数（实测 74 行 vs 真实 53 行），
> 因为那份模板的 Response/状态停留在更早时点。**只比成品 vs 成品。**

完整的本次流程记录见 [`references/20260920-FormListing处理流程总结.md`](references/20260920-FormListing处理流程总结.md)。

## 踩过的坑

| 现象 | 原因 / 解法 |
|---|---|
| `save` 抛 `PermissionError` | 输出文件被 Excel/WPS 打开，目录会生成 `~$<文件名>` 锁文件。先确认锁消失再重跑 |
| 排序结果不对 | 插入 `Asite search` 后 `Cont_Ref_No` 索引从 6 变 **7**，排序键别写错 |
| Response 切分段数偏多 | 正文自带 `" , "`，用 DP 兜底 |
| 签名比对出现"不一致" | 大小写问题（`Rex Choi` 常写成 `rex choi`），匹配前统一 lowercase |
| 表头行错位 | 源表前 6 行是导出信息，表头在**第 7 行** |
| **插入列后全表被误判为"有更新"** | `ws.insert_cols(1,3)` 若在**读取本期单元格之前**执行，按原列号 1..21 取值会整体左偏 3 列（状态列读到旗子），导致 3295 行全部标亮黄。修法：**先**在原始列位取值缓存成 `(row, coldict)`，**再** `insert_cols` |
| 上期成品列位对不上 | 上期成品可能是 21 列、也可能是已插过标记列的 24 列。用 `header_map()` **按表头名**定位列，不要写死列号 |
| UNC 路径 `FileNotFoundError` | bash 双引号里 `\\10.149.8.8\...` 会被吞成单反斜杠。改用 `//10.149.8.8/机电/...` 正斜杠形式传给 python |
| `python -c` 传入的长文本被破坏 | bash 会把反引号当命令替换（`` `文件名` `` 消失、`` `ws.insert_cols(1,3)` `` 报语法错）。长脚本/长文本一律写成文件再执行，别塞进 `-c` |
| `No module named 'openpyxl'` | 用了托管 python 3.13.12 裸解释器，换 `binaries/python/envs/default/Scripts/python.exe` |
| `ls`/`grep` 报 `command not found` | bash PATH 失效，改用绝对路径调 python，或用 PowerShell |

## 已知瑕疵（交付时如实说明）

1. 表头按用户口径写成 `MSC`，但 `Created_By` 里实际代号是 `MCS`（保持用户写法）
2. 少数行 Response 末段是孤立的 `refer to attached`（附件链接文字），按位落到了 WTP，
   语义上可能属于前一段
3. CSHK（总包）的回复/日期不落任何列，被丢弃
4. 个别行 `Cont_Ref_No` 被 Asite 误填成表单标题；排序时非 `NA[FH]` 开头的排到末尾

本次 20260920 记录还确认：4 行的 Response 末段是孤立的 `refer to attached`，可能应归入前一段；
2 行单号被误填成表单标题，另有 1 行学科段为空。这些情况保留在结果中，交付时提示人工复核。
