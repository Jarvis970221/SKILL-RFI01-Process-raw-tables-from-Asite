# -*- coding: utf-8 -*-
"""02_add_update_columns.py — 在处理过的 Form Listing 成品表 A 列前插入 3 列更新标记，
并把与上一期内容有差异的单元格标为亮黄色。

用法:
  python 02_add_update_columns.py --cur <本期成品.xlsx> --prev <上期成品.xlsx> \
      --dst <输出.xlsx> [--date 2026-09-22]

三列:
  A (RFI有更新)   : 该行有任一更新(新增/新Close/顾问新回复)时填 --date(默认当天)
  B (RFI状态更新) : 上期没有本期的单号 -> "新增"
                    本期状态=Closed 且上期不是 Closed -> "新Close"
                    两期都 Closed 但 Response_Date 有更新 -> 也填 "新Close"
  C (RFI回复更新) : Asite Reply Date (WTP/WSP/MSC) 任一列值有变化 ->
                    填 "WTP新回复"/"WSP新回复"/"MSC新回复", 多个用 "；" 连接

亮黄高亮: 上期成品按表头名逐列比对（21 个原数据列），值不同的单元格填
FFFFFF00（亮黄）。新增行整行标亮黄。表头名自动定位，兼容上期成品是否已插
入 3 列标记列。

行配对: 同一 Cont_Ref_No 多行时，先按行签名(状态+三个顾问日期+Response_Date)
配掉完全相同的行，剩余变化行按出现顺序配对（两期行顺序可能互换，不能按位配）。

固定列位(插入前的 21 列成品): 6=状态 8=Cont_Ref_No 17/18/19=Reply Date WTP/WSP/MSC 20=Response_Date
"""
import argparse, collections, datetime, os, sys
import openpyxl
from openpyxl.styles import PatternFill

HR = 7                      # 表头行
COL_STATUS, COL_REF = 6, 8
COL_DATES = {'WTP': 17, 'WSP': 18, 'MSC': 19}
COL_RESPDATE = 20
N_ORIG_COLS = 21            # 插入标记列前的成品列数
BRIGHT = PatternFill('solid', fgColor='FFFFFF00')   # 亮黄
TODAY_DEFAULT = datetime.date.today().strftime('%Y-%m-%d')
MARKERS = {'(RFI有更新)', '(RFI状态更新)', '(RFI回复更新)'}


def header_map(ws):
    """表头名 -> 列号（1-based）。跳过 3 个标记列（若上期成品已插入）。"""
    m = {}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=HR, column=c).value
        if h and str(h).strip() not in MARKERS:
            m[str(h).strip()] = c
    return m


def load_rows(path):
    """返回 {Cont_Ref_No: [rowdict, ...]}，保持出现顺序。
    rowdict: 表头名 -> 字符串值（含 __cols__ 保存表头名列表）。"""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb['Form Listing']
    hm = header_map(ws)
    out = collections.OrderedDict()
    for r in ws.iter_rows(min_row=HR + 1, values_only=True):
        ref = str(r[hm['Cont_Ref_No'] - 1] or '')
        if not ref:
            continue
        d = {name: str(r[c - 1] if c - 1 < len(r) else '' or '')
             for name, c in hm.items()}
        # 上面推导式无法区分 None 与 ''，重写一遍保证稳妥
        for name, c in hm.items():
            v = r[c - 1] if c - 1 < len(r) else None
            d[name] = '' if v is None else str(v)
        d['__cols__'] = [name for name in hm if name != 'Cont_Ref_No']
        out.setdefault(ref, []).append(d)
    wb.close()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--cur', required=True)
    ap.add_argument('--prev', required=True)
    ap.add_argument('--dst', required=True)
    ap.add_argument('--date', default=TODAY_DEFAULT)
    a = ap.parse_args()

    prev = load_rows(a.prev)
    wb = openpyxl.load_workbook(a.cur)
    ws = wb['Form Listing']

    # 本期按 key 收集（保持行序）—— 必须在 insert_cols 之前取值，列位才是原始 1..21
    cur_rows = collections.defaultdict(list)
    for r in range(HR + 1, ws.max_row + 1):
        ref = str(ws.cell(row=r, column=COL_REF).value or '')
        if not ref:
            continue
        cur = {name: str(ws.cell(row=r, column=c).value or '')
               for name, c in ((n, i) for i, n in enumerate(_COL_ORDER, 1))}
        cur_rows[ref].append((r, cur))

    ws.insert_cols(1, 3)
    heads = ['(RFI有更新)', '(RFI状态更新)', '(RFI回复更新)']
    for i, h in enumerate(heads, 1):
        ws.cell(row=HR, column=i, value=h)

    def sig(row):
        """行签名：状态 + 三个顾问日期 + Response_Date。"""
        return (row['状态'].strip(), row['Asite Reply Date (WTP)'],
                row['Asite Reply Date (WSP)'], row['Asite Reply Date (MSC)'],
                row['Response_Date'])

    n_new = n_close = n_reply = n_touch = n_cells = 0
    reply_hits = collections.Counter()
    chg_cols = collections.Counter()
    for ref, items in cur_rows.items():
        # 上期同 key 的行池（按签名消耗：先配掉完全相同的行，剩下的按序配对）
        pool = list(prev.get(ref, []))
        pairs = []
        for r, cur in items:
            hit = next((j for j, o in enumerate(pool) if sig(o) == sig(cur)), None)
            if hit is not None:
                old = pool.pop(hit)          # 完全相同的行直接配对
            else:
                old = pool.pop(0) if pool else None   # 剩余变化行按序配
            pairs.append((old, cur, r))

        for old, cur, r in pairs:
            b_val, replies = '', []
            if old is None:
                b_val = '新增'
                n_new += 1
            else:
                for t in ('WTP', 'WSP', 'MSC'):
                    if cur['Asite Reply Date (%s)' % t] != old['Asite Reply Date (%s)' % t]:
                        replies.append(t + '新回复')
                        reply_hits[t] += 1
                if cur['状态'] == 'Closed' and old['状态'] != 'Closed':
                    b_val = '新Close'
                    n_close += 1
                elif cur['状态'] == 'Closed' and old['状态'] == 'Closed' \
                        and cur['Response_Date'] != old['Response_Date']:
                    b_val = '新Close'
                    n_close += 1
            c_val = '；'.join(replies)
            if replies:
                n_reply += 1

            # 逐格比对（21 个原数据列），不同 -> 亮黄
            cols = cur.get('__cols__') or [k for k in cur if k != '__cols__']
            changed = False
            for name in cols:
                if old is None or cur[name] != old.get(name, ''):
                    ws.cell(row=r, column=cur_col(name)).fill = BRIGHT
                    changed = True
                    n_cells += 1
                    if old is not None:
                        chg_cols[name] += 1

            if b_val or c_val or changed:
                ws.cell(row=r, column=1, value=a.date)
                if b_val:
                    ws.cell(row=r, column=2, value=b_val)
                if c_val:
                    ws.cell(row=r, column=3, value=c_val)
                n_touch += 1

    ws.freeze_panes = 'D8'
    for col, w in (('A', 13), ('B', 15), ('C', 24)):
        ws.column_dimensions[col].width = w

    # 保存（处理 PermissionError 友好提示）
    try:
        wb.save(a.dst)
    except PermissionError:
        d = os.path.dirname(a.dst)
        locks = [f for f in os.listdir(d) if f.startswith('~$')]
        print('保存失败：文件被 Excel/WPS 占用。请关闭后重试。锁文件:', locks, file=sys.stderr)
        sys.exit(2)

    print('新增(上期没有): %d | 新Close: %d | 有新回复行: %d (WTP %d / WSP %d / MSC %d)'
          % (n_new, n_close, n_reply, reply_hits['WTP'], reply_hits['WSP'], reply_hits['MSC']))
    print('亮黄单元格: %d | 按列分布: %s' % (n_cells, dict(chg_cols)))
    print('共标记 %d 行 | 已保存 -> %s' % (n_touch, a.dst))


# 插入 3 列后，原第 i 列 -> i+3
_COL_ORDER = ['最后更新 (CST)', '创建日期 (CST)', '旗子', 'Asite search', '表单标题',
              '状态', '关联', 'Cont_Ref_No', 'Building-Location', 'Floor',
              'Discipline_Code', 'Response', 'Asite Reply (WTP)', 'Asite Reply (WSP)',
              'Asite Reply (MSC)', 'Created_By', 'Asite Reply Date (WTP)',
              'Asite Reply Date (WSP)', 'Asite Reply Date (MSC)', 'Response_Date',
              'Response_Flag']


def cur_col(name):
    return _COL_ORDER.index(name) + 1 + 3


if __name__ == '__main__':
    main()
