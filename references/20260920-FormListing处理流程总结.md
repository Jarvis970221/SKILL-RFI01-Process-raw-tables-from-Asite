# Form Listing 处理流程总结

> 生成日期：2026-09-20　|　脚本：`scripts/01_build_bs_listing.py`
> 源表：`C:\Users\nomo\Downloads\Form Listing.xlsx`
> 成果：`C:\Users\nomo\Downloads\20260920-处理过Form Listing.xlsx`

---

## 一、源表结构

Asite 导出的 `Form Listing.xlsx`，单工作表 `Form Listing`。

- **第 7 行才是表头**（前 6 行是导出信息/标题），数据从**第 8 行**开始
- 原始 **14 列**：

```
最后更新(CST) / 创建日期(CST) / 旗子 / 表单标题 / 状态 / 关联 /
Cont_Ref_No / Building-Location / Floor / Discipline_Code /
Response / Created_By / Response_Date / Response_Flag
```

- 数据行 **6137** 条，`Discipline_Code` 分布：`BS 3338` / `AR 1593` / `STR 1205`

---

## 二、处理步骤

### 步骤 1｜筛选专业

只保留 `Discipline_Code = BS` 的行。

```
6137 → 3338 行（删掉 AR 1593、STR 1205）
```

### 步骤 2｜剔除建筑/结构单号

`Discipline_Code` 虽标成 BS，但 `Cont_Ref_No` 学科段是建筑/结构的行也一并删掉：

| 前缀 | 删除 |
|---|---|
| `NAH-A/CS/RFI/AR/` | 16 |
| `NAH-A/CS/RFI/STR/` | 27 |

```
3338 → 3295 行
```

**保留**的其他非标准前缀（均属 BS）：

| 前缀 | 行数 |
|---|---|
| `NAH-A/CS/RFI/BS/` | 3270 |
| `NAH-A/CS/RFI/BS(AC)/` | 9 |
| `NAH-A/CS/RFI/BS(FS)/` | 9 |
| `NAH-A/CS/RFI/BS(ST)/` | 1 |
| `NAH-A/CS/CSF(SD)/BS(O)/` | 2 |
| `NAF-A/CS/RFI/BS/` | 1 |
| `NAH-A/CS/RFI//`（学科段为空） | 1 |
| 单号被误填成表单标题 | 2 |

### 步骤 3｜解析 `Created_By` → 顾问序列

`Created_By` 结构是 `人名, 公司,人名, 公司, …`（逗号分隔，**人名在前、公司名在后**）。

公司代号只有 5 种：

| 源写法 | 代号 | 角色 |
|---|---|---|
| `Wong Tung` | **WTP** | 建筑师（统筹方） |
| `WSP (Asia) Ltd` / `WSP` | **WSP** | MEP 机电顾问 |
| `MCS` | **MSC** | 幕墙 + 结构 |
| `CSCE` | CSHK | 总包（**不落列**） |

45 个人名 → 公司映射无歧义（如 `Leo Lee / Jacky Hui / Kit Lam → WSP`，
`Zoe Li / Bonnie Keung / Rex Choi → WTP`，`KC Hon / Kenneth LAW → MCS`）。

### 步骤 4｜日期按位配对

`Created_By` 中顾问出现的顺序（左→右）与 `Response_Date` 的日期（左→右）**一一对应**。

- 同一顾问出现两次 → 填两个日期，用**纯逗号**连接：`2023-09-07,2023-09-07,2023-09-11`
  （与源表 `Response_Date` 的分隔符保持一致）
- 该顾问未出现 → 留空
- CSHK（总包）的日期不落任何列

### 步骤 5｜回复按位分列

`Response` 是把多段回复用 **`" , "`（空格+逗号+空格）** 拼起来的长文本，
**段落顺序 = `Created_By` 的顾问顺序**，逐段落到对应顾问列。

**切分难点**：少数行正文里天然含 `" , "`（如 `…Block G/F , for further coordination`），
直接 split 会切多。解法：

1. 先按 `" , "` 切，段数若等于 `Created_By` 条目数 → 直接采用
2. 不等 → 用**带评分的动态规划**挑 n-1 个切点：
   - 切点后是否以 `Dear` / 人名 / `WTPL|WSP|MCS|CSHK` / `Close` 开头
   - 切点前是否以句号或人名结尾

同一顾问出现两次时，两段回复用 `\n\n----------\n\n` 连接（共 505 行）。

### 步骤 6｜排序

按 `Cont_Ref_No` **自然排序**：数字段按数值比、字母段小写比，
保证 `00042 < 00042A < 00042B < 00043`。

首条 `NAF-A/CS/RFI/BS/00102`，末条为被误填成标题的那 2 行（排到最末尾）。

### 步骤 7｜新增 `Asite search` 列

插在 `旗子` 与 `表单标题` 之间（D 列），取**表单标题开头的连续数字**（多为 5 位），
写成**字符串**并设 `number_format='@'`，保留前导 0（如 `00873`、`00108`）。

3294 行有值；1 行留空（`BS(AC)/02065` 标题开头没有数字）。

### 步骤 8｜填色 + 边框

| 颜色 | 色值 | 列 |
|---|---|---|
| **淡黄** | `#FFF2CC` | `Asite search`、`Asite Reply (WTP/WSP/MSC)`、`Asite Reply Date (WTP/WSP/MSC)` |
| **淡粉** | `#FCE4EC` | `表单标题`、`状态`、`Cont_Ref_No`、`Building-Location`、`Floor`、`Response_Date` |

共 **13 列**（淡黄 7 + 淡粉 6），表头行与全部数据行一起填色，
四边细灰边框（`thin / #808080`）。其余 8 列保持默认。

---

## 三、最终表结构（21 列 / 3295 行）

```
 1  最后更新 (CST)
 2  创建日期 (CST)
 3  旗子
 4  【Asite search】                         淡黄 · 文本格式
 5  表单标题                                 淡粉
 6  状态                                     淡粉
 7  关联
 8  Cont_Ref_No                              淡粉
 9  Building-Location                        淡粉
10  Floor                                    淡粉
11  Discipline_Code
12  Response
13  【Asite Reply (WTP)】                    淡黄
14  【Asite Reply (WSP)】                    淡黄
15  【Asite Reply (MSC)】                    淡黄
16  Created_By
17  【Asite Reply Date (WTP)】               淡黄
18  【Asite Reply Date (WSP)】               淡黄
19  【Asite Reply Date (MSC)】               淡黄
20  Response_Date                            淡粉
21  Response_Flag
```

冻结窗格 `A8`（表头行固定）。

---

## 四、结果统计

| 项目 | 数量 |
|---|---|
| 数据行 | 3295 |
| Asite search 有值 | 3294 |
| Asite Reply (WTP) | 1681 |
| Asite Reply (WSP) | 2635 |
| Asite Reply (MSC) | 204 |
| 三项回复全空（确实还没回复） | 276 |
| 含多段合并的行 | 505 |
| 含多日期的单元格 | 534 |

---

## 五、准确性验证

**签名反查交叉验证**：把每段回复末尾的署名（`Best Regards / Regards / Thanks, 人名`）
反查人名→公司，与按位归属结果比对。

> **1936 个可判定段落，100% 一致**

（注意：人名匹配必须忽略大小写，如 `Rex Choi` 常写成 `rex choi`。）

**切分正确性**：3338 行中 3327 行直接按 `" , "` 切分，段数与 `Created_By` 条目数完全相等；
剩余 10 行经动态规划切分后也全部正确。

---

## 六、已知瑕疵 / 待确认

1. **表头拼写**：按用户截图写成 `MSC`，但 `Created_By` 里实际公司代号是 `MCS`。
2. **4 行附件痕迹**：`BS/00002`、`00012`、`00009`、`00030` 的 Response 末段是孤立的
   `refer to attached`（Asite 附件链接文字），按位落到 WTP。语义上可能属于前一段 WSP。
3. **CSHK 段落不落列**：`Created_By` 中出现 `CSCE`（总包）的行，其回复/日期被丢弃。
4. **2 行单号异常**：`13902 - Query - …`、`15040 - Insufficient E&M Zone…` 的
   `Cont_Ref_No` 被误填成表单标题，已排到表末尾待人工核对。
5. **1 行学科段为空**：`NAH-A/CS/RFI//01568`。

---

## 七、踩过的坑

| 现象 | 原因 / 解法 |
|---|---|
| `openpyxl.save` 抛 `PermissionError` | 输出文件被 Excel/WPS 打开，目录会生成 `~$<文件名>` 锁文件。先确认锁消失再重跑 |
| 插入 `Asite search` 后排序失效 | 源表 `Cont_Ref_No` 是索引 6，插入新列后变成 **7**，排序键别写错 |
| `Response` 切分段数多于期望 | 正文里本身含 `" , "`，改用带评分的动态规划挑切点 |
| 签名比对出现"不一致" | 大小写问题，匹配前统一 `.lower()` |
