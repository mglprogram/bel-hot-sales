# -*- coding: utf-8 -*-
"""
十部半托管-库存管理表 -> 热销销量表（V3 格式）

用法:
    python bel_pipeline.py "<十部半托管-库存管理表.xlsx>" [输出路径] [--brand BEL]
                                                          [--last-week 9.12-9.18]
                                                          [--prev-week 9.05-9.11]

规则:
  1. 数据源 sheet: 总销售表
  2. 筛选: MSKU 第2段(品牌段) == --brand   （该表没有以品牌名开头的 MSKU）
  3. 聚合粒度: 款号
  4. 日期窗口: 默认自动取源表「最后一个有数据的日期列」为 D0
       上周   = D0-6 .. D0      (7天, 拆成每日列)
       上上周 = D0-13 .. D0-7   (7天, 汇总一列)
     也可用 --last-week / --prev-week 显式指定(格式 M.D-M.D)
  5. 保留条件: 上上周合计 > 0 且 上周合计 > 0   （单周有销量的进"已排除"表）
  6. 排序: 上周合计 desc, 再 上上周合计 desc, 再 款号
  7. 输出: 两个 sheet，"两周均有销量" + "已排除-仅单周有销量"，V3 视觉规格

注意(踩过的坑):
  * MSKU 形如 TM-BEL-102503-5SetA-S，品牌在第2段，不是前缀。
  * 表头存在重复日期标签(如 4.24日 出现两次)，因此一律按"列位置"处理，
    绝不能用 (月,日) 做 dict key，否则窗口会错位。
"""
import sys, re, io, argparse, os
from collections import OrderedDict
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Windows 控制台默认 cp936，中文输出会变成乱码；统一切到 UTF-8。
# 这只影响 stdout 显示，不影响写出的 xlsx 内容。
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SRC_SHEET = "总销售表"
COL_MSKU, COL_STYLE, COL_SPU, COL_DA7 = 11, 7, 9, 13   # 1-based

# ---- V3 style spec ----
THIN = Side(style="thin", color="BFBFBF")
BD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_TOP3   = PatternFill("solid", fgColor="FFFFF2CC")
F_DAILY  = PatternFill("solid", fgColor="FFFCE4D6")
F_BLUE   = PatternFill("solid", fgColor="FF4472C4")
F_ORANGE = PatternFill("solid", fgColor="FFED7D31")
C_UP, C_DOWN, C_FLAT = "FFC00000", "FF0070C0", "FF808080"
FONT_SZ = 11
C_LAST, C_DA7_, C_WOW = 11, 12, 13
L1, L7 = 4, 10
NCOL = 13


def d(i):
    """日期显示: (9,5) -> '9.05'"""
    return "%d.%02d" % i


def num(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(',', '')
    if not s: return 0.0
    try: return float(s)
    except ValueError: return 0.0


def st(v):
    return '' if v is None else str(v).strip()


def parse_span(txt):
    """'9.12-9.18' -> ((9,12),(9,18))"""
    m = re.fullmatch(r"\s*(\d+)\.(\d+)\s*-\s*(\d+)\.(\d+)\s*", txt)
    if not m:
        raise SystemExit("窗口格式应为 M.D-M.D，例如 9.12-9.18，收到: %r" % txt)
    return (int(m.group(1)), int(m.group(2))), (int(m.group(3)), int(m.group(4)))


def all_date_cols(hdr):
    """返回 [(col_idx, (mo,dd), label)]，按表头位置顺序"""
    cols = []
    for i, h in enumerate(hdr):
        hs = "" if h is None else str(h).strip()
        m = re.fullmatch(r"(\d+)\.(\d+)日", hs)
        if m:
            cols.append((i, (int(m.group(1)), int(m.group(2))), hs))
    return cols


def pick_span(cols, a, b):
    """在 cols 里按 (mo,dd) 取一个闭区间。

    表头里同一日期可能**重复出现**（本表 9.5..9.18 在 col128-141 与 col860-873
    各出现一次；早月还有 4.24日 连续两列）。因此必须从**每一组**候选里选，
    而不能盲目取第一次出现，否则会把窗口定位到错误的区块。
    判定标准：区间天数正好等于 (b - a) 的日历天数 + 1，且起点/终点匹配。
    """
    want = (b[0] - a[0]) * 31 + (b[1] - a[1]) + 1   # 宽松天数（同月时即真实天数）
    cands = []
    starts = [k for k, c in enumerate(cols) if c[1] == a]
    for start in starts:
        for k in range(start, min(start + 40, len(cols))):
            if cols[k][1] == b:
                span = cols[start:k + 1]
                cands.append((start, span))
                break
    if cands:
        # 优先选长度最接近期望的；并列时取位置最靠后的（后区通常是当期数据区）
        cands.sort(key=lambda t: (abs(len(t[1]) - want), -t[0]))
        return cands[0][1]
    # 退而求其次：取最后一组 a 起连续 7 列
    if starts:
        span = cols[starts[-1]:starts[-1] + 7]
        if len(span) == 7:
            print("!! 提示: 未能按终点 %s 精确定位，改用 %s 起连续 7 列"
                  % (d(b), d(a)))
            return span
    raise SystemExit("窗口 %s-%s 无法定位" % (d(a), d(b)))


def detect_windows(hdr, ws_rows):
    """默认: 最后一个有数据的日期列往前 14 列 -> (上周7列, 上上周7列, D0)

    表头里同一日期可能出现在多个区块（本表 9.5..9.18 在 col128-141 与 col860-873
    各出现一次，早区块是残留的稀疏数据）。这里以「最后一个有数据的日期列」为锚点
    向前取连续 14 列，天然落在当期数据块里，不受重复区块影响。
    同时检查两个区块的数据密度，若差异悬殊则提示，避免用户误以为数据缺失。
    """
    cols = all_date_cols(hdr)
    if not cols:
        raise SystemExit("总销售表里找不到日期列（形如 9.18日）")
    counts = [0] * len(cols)
    for row in ws_rows:
        for k, (i, _k, _lbl) in enumerate(cols):
            v = row[i]
            if v is not None and str(v).strip() != "":
                counts[k] += 1
    live = [k for k in range(len(cols)) if counts[k] > 0]
    if not live:
        raise SystemExit("总销售表里找不到任何有数据的日期列")
    end = live[-1]
    start = max(0, end - 13)
    win = cols[start:end + 1]
    if len(win) < 14:
        print("!! 警告: 可用日期列只有 %d 个(不足14)，两侧窗口会不对称" % len(win))

    # 重复日期区块检测：同一 (月,日) 若在别处也出现，提示其密度差异
    dupes = []
    for k in range(len(cols)):
        if k in range(start, end + 1):
            continue
        for kk in range(start, end + 1):
            if cols[k][1] == cols[kk][1]:
                dupes.append((cols[k][1], cols[kk][0] + 1, cols[k][0] + 1,
                              counts[kk], counts[k]))
                break
    if dupes:
        mm, live_col, old_col, live_n, old_n = dupes[0]
        print("!! 注意: 日期 %s 在表头出现多次 —— 本脚本取的是 col%d(有数据 %d 行)，"
              "另有 col%d(有数据 %d 行) 是残留区块，已忽略。"
              % (d(mm), live_col, live_n, old_col, old_n))
        print("         按列位置定位可以避开这个坑；切勿按 (月,日) 建字典，会取错区块。")

    # 窗口连续性检查：14 个日期列应当是逐日递增的
    seq = [c[1] for c in win]
    for j in range(1, len(seq)):
        pm, pd = seq[j - 1]
        cm, cd = seq[j]
        expect = (pm, pd + 1) if pd < 30 else (pm + 1, 1)
        if (cm, cd) != expect:
            print("!! 警告: 窗口内日期不连续: %s -> %s" % (d(seq[j-1]), d(seq[j])))

    prev7 = win[:7]
    last7 = win[7:14]
    return last7, prev7, cols[end][1]


def build(sh, data, title, subtitle):
    sh.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOL)
    c = sh.cell(row=1, column=1, value=title)
    c.font = Font(bold=True, size=13)
    c.alignment = Alignment(horizontal="center", vertical="center")
    sh.row_dimensions[1].height = 26

    sh.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOL)
    c = sh.cell(row=2, column=1, value=subtitle)
    c.font = Font(size=9, italic=True, color="FF595959")
    c.alignment = Alignment(horizontal="center")

    sh.merge_cells(start_row=3, start_column=L1, end_row=3, end_column=L7)
    cc = sh.cell(row=3, column=L1, value=sh._banner)
    cc.font = Font(bold=True, color="FFFFFFFF", size=10)
    cc.fill = F_ORANGE
    cc.alignment = Alignment(horizontal="center", vertical="center")

    sh.row_dimensions[4].height = 30
    for j, h in enumerate(sh._headers, 1):
        cell = sh.cell(row=4, column=j, value=h)
        cell.font = Font(bold=True, size=FONT_SZ, color="FFFFFFFF")
        cell.fill = F_BLUE
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, a in enumerate(data, 1):
        r = 4 + i
        wow = (a['last7'] - a['prev7']) / a['prev7'] if a['prev7'] else None
        vals = ([i, a['style'], round(a['prev7'], 1)] +
                [round(x, 1) for x in a['l']] +
                [round(a['last7'], 1), round(a['da7'], 2), wow])
        top3 = i <= 3
        for j, v in enumerate(vals, 1):
            cell = sh.cell(row=r, column=j, value=v)
            cell.border = BD
            if j == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif j == C_WOW:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif j >= 3:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            if j == C_WOW:
                cell.number_format = "+0.0%;-0.0%;0.0%"
                cell.font = Font(bold=True, size=FONT_SZ,
                                 color=(C_UP if (v or 0) > 0 else C_DOWN if (v or 0) < 0 else C_FLAT))
            elif j >= 3:
                cell.number_format = "#,##0.0"
                cell.font = Font(size=FONT_SZ, bold=(j == C_LAST))
            elif j == 1:
                cell.font = Font(size=FONT_SZ, bold=top3)
            else:
                cell.font = Font(size=FONT_SZ)
            if L1 <= j <= L7:
                cell.fill = F_DAILY
            elif top3 and j != C_WOW:
                cell.fill = F_TOP3

    sh.column_dimensions['A'].width = 5.5
    sh.column_dimensions['B'].width = 17
    sh.column_dimensions['C'].width = 15
    for k in range(7):
        sh.column_dimensions[get_column_letter(L1 + k)].width = 7.2
    sh.column_dimensions[get_column_letter(C_LAST)].width = 15
    sh.column_dimensions[get_column_letter(C_DA7_)].width = 13
    sh.column_dimensions[get_column_letter(C_WOW)].width = 15
    sh.freeze_panes = "C5"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst", nargs="?", default=None)
    ap.add_argument("--brand", default="BEL")
    ap.add_argument("--last-week", default=None, help="M.D-M.D")
    ap.add_argument("--prev-week", default=None, help="M.D-M.D")
    args = ap.parse_args()
    brand = args.brand.strip().upper()
    dst = args.dst or ("%s两周有销量_输出.xlsx" % brand)

    wb = openpyxl.load_workbook(args.src, read_only=True, data_only=True)
    if SRC_SHEET in wb.sheetnames:
        wsname = SRC_SHEET
    else:
        cand = [s for s in wb.sheetnames if "销售表" in s]
        if not cand:
            raise SystemExit("找不到 总销售表 / *销售表 sheet: %s" % wb.sheetnames)
        wsname = cand[0]
    ws = wb[wsname]
    it = ws.iter_rows(min_row=1, values_only=True)
    hdr = next(it)
    all_rows = list(it)
    print("来源 sheet: %s  数据行: %d" % (wsname, len(all_rows)))

    cols = all_date_cols(hdr)
    if args.last_week or args.prev_week:
        if not (args.last_week and args.prev_week):
            raise SystemExit("--last-week 与 --prev-week 必须同时给出")
        a1, a2 = parse_span(args.last_week)
        b1, b2 = parse_span(args.prev_week)
        last7 = pick_span(cols, a1, a2)
        prev7 = pick_span(cols, b1, b2)
        d0 = last7[-1][1]
        print("D0 = %s (显式指定)" % d(last7[-1][1]))
    else:
        last7, prev7, d0 = detect_windows(hdr, all_rows)
        print("D0(最后有数据日) = %s" % d(d0))

    print("上周窗口   = %s .. %s" % (d(last7[0][1]), d(last7[-1][1])))
    print("上上周窗口 = %s .. %s" % (d(prev7[0][1]), d(prev7[-1][1])))
    if len(last7) != 7 or len(prev7) != 7:
        print("!! 警告: 窗口天数 上周=%d 上上周=%d (正常情况下各为7)" % (len(last7), len(prev7)))

    L_IDX = [c[0] for c in last7]
    P_IDX = [c[0] for c in prev7]
    MI, SI, SPI, DI = COL_MSKU - 1, COL_STYLE - 1, COL_SPU - 1, COL_DA7 - 1

    agg = OrderedDict()
    brand_rows = 0
    for row in all_rows:
        ms = st(row[MI])
        parts = ms.split('-')
        if len(parts) < 2 or parts[1].upper() != brand:
            continue
        brand_rows += 1
        key = st(row[SI]) or ms
        a = agg.setdefault(key, {'style': key, 'prev7': 0.0,
                                 'l': [0.0] * len(L_IDX), 'da7': 0.0, 'spu': set()})
        a['spu'].add(st(row[SPI]))
        a['da7'] += num(row[DI])
        for i in P_IDX: a['prev7'] += num(row[i])
        for k, i in enumerate(L_IDX): a['l'][k] += num(row[i])
    for a in agg.values():
        a['last7'] = sum(a['l'])

    if not agg:
        raise SystemExit(
            "没有匹配品牌 %s 的 MSKU。请确认品牌代号；"
            "注意该表 MSKU 格式为 平台-品牌-款号，品牌在第2段。" % brand)

    keep = [a for a in agg.values() if a['prev7'] > 0 and a['last7'] > 0]
    keep.sort(key=lambda x: (-x['last7'], -x['prev7'], x['style']))
    excl = [a for a in agg.values() if (a['prev7'] > 0) != (a['last7'] > 0)]
    excl.sort(key=lambda x: -(x['prev7'] + x['last7']))

    print("%s 行数=%d  款号=%d  保留=%d  排除=%d" %
          (brand, brand_rows, len(agg), len(keep), len(excl)))

    last_days = [c[2] for c in last7]
    pv_lbl = "%s-%s" % (d(prev7[0][1]), d(prev7[-1][1]))
    ls_lbl = "%s-%s" % (d(last7[0][1]), d(last7[-1][1]))
    pv_en = pv_lbl.replace("-", "\u2013")
    ls_en = ls_lbl.replace("-", "\u2013")

    out = openpyxl.Workbook()
    sh = out.active
    sh.title = "两周均有销量"
    sh._banner = "上周每日销量（%s - %s）" % (d(last7[0][1]), d(last7[-1][1]))
    sh._headers = (["排名", "款号", "上上周销量合计(%s)" % pv_lbl] + last_days +
                   ["上周销量合计(%s)" % ls_lbl, "最近7天日均", "上周环比上上周"])
    build(sh, keep,
          "%s 品牌 · 两周均有销量产品（%d 款）" % (brand, len(keep)),
          "来源：%s ｜ 筛选：MSKU 品牌段=%s 且 上上周(%s)>0 且 上周(%s)>0 ｜ 按款号聚合 ｜ 上上周为合计，上周按天拆分"
          % (wsname, brand, pv_en, ls_en))

    sh2 = out.create_sheet("已排除-仅单周有销量")
    H2 = ["款号", "上上周合计(%s)" % pv_lbl, "上周合计(%s)" % ls_lbl, "排除原因"]
    sh2.row_dimensions[1].height = 30
    for j, h in enumerate(H2, 1):
        cell = sh2.cell(row=1, column=j, value=h)
        cell.font = Font(bold=True, size=FONT_SZ, color="FFFFFFFF")
        cell.fill = F_BLUE
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, a in enumerate(excl, 1):
        why = "仅上上周有销量" if a['last7'] == 0 else "仅上周有销量"
        for j, v in enumerate([a['style'], a['prev7'], a['last7'], why], 1):
            cell = sh2.cell(row=1 + i, column=j, value=v)
            cell.border = BD
            cell.font = Font(size=FONT_SZ)
            if j == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif j == 4:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0.0"
    for col, w in zip("ABCD", [18, 15, 15, 18]):
        sh2.column_dimensions[col].width = w
    sh2.freeze_panes = "A2"

    out.save(dst)
    print("已保存:", dst)

    # 自检: 每日之和必须等于上周合计
    bad = 0
    for r in range(5, sh.max_row + 1):
        v = [sh.cell(row=r, column=c).value for c in range(1, NCOL + 1)]
        if round(sum(v[3:10]), 1) != round(v[10], 1):
            bad += 1
    print("校验: 日列之和 vs 上周合计 不符行数 = %d" % bad)


if __name__ == "__main__":
    main()
