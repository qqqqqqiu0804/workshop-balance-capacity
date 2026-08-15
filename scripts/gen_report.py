# -*- coding: utf-8 -*-
"""
车间月产量平衡点（B）分析报告生成引擎
====================================
读生产订单明细 xlsx -> 实时聚合月产量 -> 两套算法算候选平衡点 B ->
渲染单文件老板版 HTML 报告（内联 CSS + 内联 SVG，无外部依赖）。

用法:
    python gen_report.py                         # 用 CONFIG 默认值（相对当前目录的 订单明细.xlsx）
    python gen_report.py <源.xlsx> <输出.html>    # 命令行覆盖路径

所有口径开关都在 CONFIG 里；逻辑无需改动即可复用。
"""
import argparse
import os
import sys
import openpyxl
import statistics
from collections import defaultdict
from datetime import datetime

# ============================ CONFIG（按需修改） ============================
CONFIG = {
    # 数据源 / 输出（命令行可覆盖）
    # 默认相对当前工作目录；公开仓库不含任何个人绝对路径。
    # 实际使用时请以 CLI 传参为准：python gen_report.py <源.xlsx> <输出.html>
    "src": "订单明细.xlsx",
    "out": "车间平衡点产能分析报告_老板版.html",
    # 订单明细 sheet 名与列下标（0 起）
    "sheet": "订单明细",
    "date_col": 0,      # 订单日期
    "qty_col": 7,       # 数量（个）= 产量
    # 月份范围（建模口径）
    "all_months": ["2026-01", "2026-02", "2026-03", "2026-04",
                   "2026-05", "2026-06", "2026-07", "2026-08"],
    "complete_months": {"2026-03", "2026-04", "2026-05", "2026-06", "2026-07"},
    # 工作日标准化用的标准工作天数
    "std_days": 26,
    # 报告元信息
    "report_title": "车间月产量平衡点（B）分析报告",
    "report_gen_date": "2026-08-15",
}
# ==========================================================================

def load(src, sheet, date_col, qty_col):
    wb = openpyxl.load_workbook(src, data_only=True)
    od = wb[sheet]
    rows = list(od.iter_rows(values_only=True))
    data = rows[1:]
    qty_by_month = defaultdict(float)
    days_by_month = defaultdict(set)
    order_no_empty = 0
    status_filled = 0
    for r in data:
        d = r[date_col]
        qty = r[qty_col]
        ono = r[2] if len(r) > 2 else None
        st = r[9] if len(r) > 9 else None
        if d is None:
            continue
        dt = d if hasattr(d, "year") else None
        if dt is None:
            try:
                dt = datetime.strptime(str(d)[:10], "%Y-%m-%d")
            except Exception:
                continue
        m = dt.strftime("%Y-%m")
        days_by_month[m].add(dt.day)
        try:
            if qty is not None:
                qty_by_month[m] += float(qty)
        except Exception:
            pass
        if ono in (None, ""):
            order_no_empty += 1
        if st not in (None, ""):
            status_filled += 1
    return qty_by_month, days_by_month, order_no_empty, status_filled


def stats(x):
    return (round(statistics.mean(x)), round(statistics.median(x)),
            round(statistics.mean(x) * 0.95), round(max(x) * 0.80))


def fmt(n):
    return f"{n:,}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="?", default=CONFIG["src"])
    ap.add_argument("out", nargs="?", default=CONFIG["out"])
    args = ap.parse_args()
    SRC, OUT = args.src, args.out
    C = CONFIG

    if not os.path.exists(SRC):
        print(f"[错误] 找不到数据源：{SRC}")
        print(f"请通过命令行指定：python gen_report.py <源.xlsx> <输出.html>")
        print(f"或直接把源文件放到当前目录并命名为：{CONFIG['src']}")
        sys.exit(1)

    qty_by_month, days_by_month, order_no_empty, status_filled = load(
        SRC, C["sheet"], C["date_col"], C["qty_col"])
    order = C["all_months"]
    complete = C["complete_months"]
    cm = [m for m in order if m in complete]
    vals = [qty_by_month[m] for m in cm]
    wdays = [len(days_by_month[m]) for m in cm]
    daily = [v / d for v, d in zip(vals, wdays)]
    STD = C["std_days"]
    norm = [d * STD for d in daily]

    raw_mean, raw_med, raw_cons, raw_peak = stats(vals)
    norm_mean, norm_med, norm_cons, norm_peak = stats(norm)
    total_cm = sum(vals)
    peak_m = max(vals)
    peak_label = cm[vals.index(peak_m)]

    candidates = [
        ("历史原始·均值", raw_mean, "完整月产量求平均（受低值月拖累，偏松）"),
        ("历史原始·中位数", raw_med, "完整月产量取中位（抗异常，推荐基准）"),
        ("历史原始·保守(均值×0.95)", raw_cons, "均值下浮5%，提成更克制"),
        ("历史原始·峰值×80%", raw_peak, "业内标准定额参考：预计峰值×80%"),
        ("工作日标准化·均值", norm_mean, f"先算日均产能×{STD}标准天，再平均"),
        ("工作日标准化·中位数", norm_med, f"先折算标准天取中位（剔除各月天数干扰，更纯产能）"),
        ("工作日标准化·保守(均值×0.95)", norm_cons, f"标准化均值下浮5%"),
        ("工作日标准化·峰值×80%", norm_peak, f"标准化峰值×80%"),
    ]
    sens = []
    for name, b, _ in candidates:
        om = sum(1 for v in vals if v > b)
        ex = sum(max(0, v - b) for v in vals)
        sens.append((name, b, om, int(ex)))

    # 月度表
    month_rows = ""
    for m in order:
        q = int(qty_by_month.get(m, 0))
        dd = sorted(days_by_month.get(m, []))
        nd = len(dd)
        span = f"{min(dd)}–{max(dd)}" if dd else "—"
        is_c = m in complete
        tag = '<span class="tag tag-green">✓ 完整月</span>' if is_c else '<span class="tag tag-yellow">✗ 不完整</span>'
        month_rows += (f"<tr><td>{m}</td><td class='text-right'>{fmt(q)}</td>"
                       f"<td class='text-right'>{nd}</td><td>{span}</td><td>{tag}</td></tr>")

    kpi = f"""
<div class="kpi-card"><div class="label">数据覆盖</div><div class="value blue">1–8月</div><div class="unit">共 8 个自然月</div></div>
<div class="kpi-card"><div class="label">完整运营月</div><div class="value green">5 个</div><div class="unit">3–7 月（剔除1/2/8）</div></div>
<div class="kpi-card"><div class="label">完整月总产量</div><div class="value orange">{fmt(total_cm)}</div><div class="unit">个</div></div>
<div class="kpi-card"><div class="label">月均产量</div><div class="value">{fmt(raw_mean)}</div><div class="unit">个/月（均值）</div></div>
<div class="kpi-card"><div class="label">月产量中位数</div><div class="value green">{fmt(raw_med)}</div><div class="unit">个/月（抗异常）</div></div>
<div class="kpi-card"><div class="label">单月峰值</div><div class="value red">{fmt(peak_m)}</div><div class="unit">个（{peak_label[5:]}月）</div></div>
"""

    cand_tbl = "".join(
        f"<tr><td>{n}</td><td class='text-right'><b>{fmt(b)}</b></td><td class='text-dim'>{desc}</td></tr>"
        for n, b, desc in candidates)
    sens_tbl = "".join(
        f"<tr><td>{n}</td><td class='text-right'>{fmt(b)}</td><td class='text-right'>{om}/5</td><td class='text-right'>{fmt(ex)}</td></tr>"
        for n, b, om, ex in sens)

    # 柱状图 SVG
    ch = 220; base = ch - 34; x0 = 56; bw = 64; gp = 26
    cmax = max(vals + [norm_med, raw_med])
    svg_bars = ""
    for i, (m, v) in enumerate(zip(cm, vals)):
        x = x0 + i * (bw + gp)
        h = v / cmax * (base - 10)
        svg_bars += f'<rect x="{x}" y="{base-h:.0f}" width="{bw}" height="{h:.0f}" rx="3" fill="url(#g1)"></rect>'
        svg_bars += f'<text x="{x+bw/2}" y="{base-h-8:.0f}" font-size="11" text-anchor="middle" fill="#cfd6e4">{v/10000:.1f}万</text>'
        svg_bars += f'<text x="{x+bw/2}" y="{base+16}" font-size="11" text-anchor="middle" fill="#9ca3af">{m[5:]}月</text>'

    def line(bv, col, lab):
        y = base - (bv / cmax * (base - 10))
        return (f'<line x1="{x0-8}" y1="{y:.0f}" x2="{x0+5*(bw+gp)-8}" y2="{y:.0f}" '
                f'stroke="{col}" stroke-width="2" stroke-dasharray="6 4"></line>'
                f'<text x="{x0+5*(bw+gp)-4}" y="{y-5:.0f}" font-size="10" fill="{col}" text-anchor="end">{lab} {bv/10000:.1f}万</text>')
    svg = (f'<svg viewBox="0 0 {x0+5*(bw+gp)+20} {ch}" width="100%" style="max-width:620px">'
           f'<defs><linearGradient id="g1" x1="0" y1="0" x2="0" y2="1">'
           f'<stop offset="0%" stop-color="#e85d3a"/><stop offset="100%" stop-color="#e85d3a66"/></linearGradient></defs>'
           f'{svg_bars}{line(raw_med,"#3ecf8e","原始中位")}{line(norm_med,"#f0a030","标准化中位")}'
           f'<line x1="{x0-8}" y1="{base}" x2="{x0+5*(bw+gp)-8}" y2="{base}" stroke="#2d323d"></line>'
           f'</svg>')

    methods = f"""
<div class="method-card m1">
  <div class="method-header"><span class="method-name">方法一：历史原始中位数法（推荐主方法）</span><span class="method-tag tag-green">数据驱动·抗异常</span></div>
  <div class="method-body"><p><strong>原理</strong>：取完整月产量的中位数作为平衡点。中位数天然抗极端值。</p>
  <div class="formula">平衡点月产能 = Median(完整月产量序列)</div></div>
  <div class="method-result"><strong>测算结果</strong>：B = {fmt(raw_med)} 个/月。以 4月、5月 超产，3/6/7月 不超，区分度合理。</div>
</div>
<div class="method-card m2">
  <div class="method-header"><span class="method-name">方法二：工作日标准化法（更纯产能）</span><span class="method-tag tag-blue">剔除天数干扰</span></div>
  <div class="method-body"><p><strong>原理</strong>：各月实际工作天数不同，先折算"日均产能×{STD}标准天"再取中位，剔除"少开工"对基线的污染。</p>
  <div class="formula">日均产能 = 月产量 ÷ 实际工作天数<br>平衡点月产能 = Median(日均产能) × {STD} 标准天</div></div>
  <div class="method-result"><strong>测算结果</strong>：B = {fmt(norm_med)} 个/月。比原始法高，因为低产月多为"开工天数少"而非"产能低"。</div>
</div>
<div class="method-card m3">
  <div class="method-header"><span class="method-name">方法三：保守 / 比例法（业内参考）</span><span class="method-tag tag-yellow">稳健校准</span></div>
  <div class="method-body"><p><strong>原理</strong>：用均值下浮 5% 作保守基线，或用"预计峰值×80%"作标准定额参考。</p>
  <div class="formula">保守 B = 均值 × 0.95 = {fmt(raw_cons)}<br>定额 B = 峰值 × 80% = {fmt(raw_peak)}</div></div>
  <div class="method-result"><strong>定位</strong>：作为主方法的下界校准，防止基线定得偏高导致提成过松。</div>
</div>
<div class="method-card m4">
  <div class="method-header"><span class="method-name">方法四：成本覆盖 / 理论产能法（升级路径）</span><span class="method-tag tag-orange">需补充数据</span></div>
  <div class="method-body"><p><strong>原理</strong>：从财务（固定成本÷单位边际）或设备（理论产能×负荷率）自上而下算"应达产能"。</p>
  <div class="formula">成本 B = 月固定成本 ÷ 单位边际<br>理论 B = 机台理论产能 × 有效工时 × 目标达成率</div></div>
  <div class="method-result"><strong>当前状态</strong>：⚠️ 缺成本/单价/标准工时数据，暂不可算，留作补数据后交叉验证。</div>
</div>
"""

    timeline = f"""
<div class="timeline">
  <div class="timeline-item"><div class="time">启动期</div><div class="title">选定初值</div><div class="desc">管理层从本报告候选区间（{fmt(raw_med)}–{fmt(norm_med)}）中选定一个 B 作为首月基准。</div></div>
  <div class="timeline-item"><div class="time">第 2 月起 · 自动</div><div class="title">滚动更新</div><div class="desc">每月 B = 前 5 个完整月产量的中位数，自动随产能趋势漂移，无需每次开会重拍。</div></div>
  <div class="timeline-item"><div class="time">结构性变化</div><div class="title">人工覆盖</div><div class="desc">遇工艺升级 / 新设备 / 产线调整，管理层可指定某段时期 B，覆盖滚动值。</div></div>
  <div class="timeline-item"><div class="time">异常事件</div><div class="title">窗口剔除</div><div class="desc">已知一次性事件月可在滚动窗口标记排除，不参与基线。</div></div>
  <div class="timeline-item"><div class="time">防护</div><div class="title">下限保护</div><div class="desc">B 不低于历史最低完整月产量的某比例（如 90%），防止连续低月拖垮基线。</div></div>
</div>
"""

    issues = f"""
<li><span class="priority tag-red">P0</span><div class="content"><strong>订单号整列空</strong><span>约 {order_no_empty} 行订单号全空，无法做订单级去重与对账，须补或明确弃用。</span></div></li>
<li><span class="priority tag-red">P0</span><div class="content"><strong>订单状态几乎未填</strong><span>仅 {status_filled} 行填"完成"，产量"完成/在制"无法区分，须规范填报。</span></div></li>
<li><span class="priority tag-yellow">P2</span><div class="content"><strong>6 月产量异常低</strong><span>265,556 个仅 21 个工作日记录，须确认是否一次性事件（检修/订单缺口）。</span></div></li>
"""

    CSS = """<style>
:root{--bg:#0f1115;--card:#1a1d24;--card-light:#222630;--border:#2d323d;--accent:#e85d3a;--accent-2:#f0a030;--accent-3:#3ecf8e;--accent-4:#6c8eef;--text:#e8e8e8;--text-dim:#9ca3af;--text-mute:#6b7280;--red:#ef4444;--green:#22c55e;--yellow:#eab308;--orange:#f97316;}
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:-apple-system,"Microsoft YaHei","PingFang SC",sans-serif;background:var(--bg);color:var(--text);line-height:1.7;padding:0;}
.container{max-width:960px;margin:0 auto;padding:40px 24px 80px;}
.report-header{text-align:center;padding:60px 24px 48px;background:linear-gradient(135deg,#1a1d24 0%,#252a35 100%);border-bottom:1px solid var(--border);border-radius:0 0 24px 24px;margin-bottom:48px;}
.report-header .badge{display:inline-block;padding:4px 14px;background:rgba(232,93,58,0.15);border:1px solid rgba(232,93,58,0.3);border-radius:20px;font-size:12px;color:var(--accent);margin-bottom:20px;letter-spacing:1px;}
.report-header h1{font-size:32px;font-weight:700;margin-bottom:12px;background:linear-gradient(135deg,#e85d3a,#f0a030);-webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.report-header .subtitle{font-size:15px;color:var(--text-dim);max-width:640px;margin:0 auto;}
.report-header .meta{margin-top:24px;display:flex;justify-content:center;flex-wrap:wrap;gap:24px;font-size:13px;color:var(--text-mute);}
.report-header .meta span{display:flex;align-items:center;gap:6px;}
.section{margin-bottom:48px;}
.section-title{font-size:22px;font-weight:700;margin-bottom:8px;padding-left:14px;border-left:4px solid var(--accent);display:flex;align-items:center;gap:10px;}
.section-title .num{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;background:var(--accent);color:#fff;border-radius:6px;font-size:14px;font-weight:700;}
.section-desc{font-size:14px;color:var(--text-dim);margin-bottom:20px;padding-left:18px;}
.card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:24px;margin-bottom:16px;}
.card-title{font-size:16px;font-weight:600;margin-bottom:12px;color:var(--accent-2);display:flex;align-items:center;gap:8px;}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:24px;}
.kpi-card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:20px;text-align:center;}
.kpi-card .label{font-size:12px;color:var(--text-dim);margin-bottom:8px;}
.kpi-card .value{font-size:26px;font-weight:700;}
.kpi-card .value.red{color:var(--red);}.kpi-card .value.green{color:var(--green);}.kpi-card .value.orange{color:var(--orange);}.kpi-card .value.blue{color:var(--accent-4);}
.kpi-card .unit{font-size:12px;color:var(--text-dim);margin-top:4px;}
table{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:8px;}
th{background:var(--card-light);color:var(--text-dim);font-weight:600;padding:10px 12px;text-align:left;border-bottom:1px solid var(--border);white-space:nowrap;}
td{padding:10px 12px;border-bottom:1px solid var(--border);}
tr:hover td{background:rgba(255,255,255,0.02);}
.text-right{text-align:right;}.text-center{text-align:center;}.text-dim{color:var(--text-dim);}
.tag{display:inline-block;padding:2px 10px;border-radius:4px;font-size:11px;font-weight:600;}
.tag-red{background:rgba(239,68,68,0.15);color:var(--red);}.tag-green{background:rgba(34,197,94,0.15);color:var(--green);}
.tag-yellow{background:rgba(234,179,8,0.15);color:var(--yellow);}.tag-blue{background:rgba(108,142,239,0.15);color:var(--accent-4);}.tag-orange{background:rgba(249,115,22,0.15);color:var(--orange);}
.alert{border-radius:10px;padding:16px 20px;margin-bottom:12px;border-left:4px solid;font-size:14px;}
.alert-warning{background:rgba(234,179,8,0.08);border-color:var(--yellow);}
.alert-danger{background:rgba(239,68,68,0.08);border-color:var(--red);}
.alert-info{background:rgba(108,142,239,0.08);border-color:var(--accent-4);}
.alert-success{background:rgba(34,197,94,0.08);border-color:var(--green);}
.alert strong{display:block;margin-bottom:4px;}
.method-card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:24px;margin-bottom:16px;position:relative;overflow:hidden;}
.method-card::before{content:'';position:absolute;left:0;top:0;bottom:0;width:4px;}
.method-card.m1::before{background:var(--accent-3);}.method-card.m2::before{background:var(--accent-4);}.method-card.m3::before{background:var(--accent-2);}.method-card.m4::before{background:var(--accent);}
.method-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:8px;}
.method-name{font-size:16px;font-weight:600;}.method-tag{font-size:11px;padding:3px 10px;border-radius:4px;font-weight:600;}
.method-body{font-size:14px;color:var(--text-dim);}.method-body p{margin-bottom:8px;}
.method-result{margin-top:12px;padding:12px 16px;background:var(--card-light);border-radius:8px;font-size:13px;}
.formula{background:var(--card-light);border:1px solid var(--border);border-radius:8px;padding:16px;margin:12px 0;font-family:"Consolas","Courier New",monospace;font-size:14px;color:var(--accent-2);text-align:center;line-height:2;}
.timeline{position:relative;padding-left:28px;}.timeline::before{content:'';position:absolute;left:8px;top:0;bottom:0;width:2px;background:var(--border);}
.timeline-item{position:relative;margin-bottom:24px;padding-left:8px;}
.timeline-item::before{content:'';position:absolute;left:-24px;top:6px;width:12px;height:12px;border-radius:50%;background:var(--card);border:2px solid var(--accent);}
.timeline-item .time{font-size:12px;color:var(--accent);font-weight:600;margin-bottom:4px;}
.timeline-item .title{font-size:15px;font-weight:600;margin-bottom:4px;}
.timeline-item .desc{font-size:13px;color:var(--text-dim);}
.gap-list{list-style:none;}.gap-list li{padding:10px 0;border-bottom:1px solid var(--border);display:flex;align-items:flex-start;gap:10px;font-size:14px;}
.gap-list li:last-child{border-bottom:none;}
.gap-list .priority{flex-shrink:0;width:50px;text-align:center;font-size:11px;font-weight:700;padding:2px 6px;border-radius:4px;}
.gap-list .content{flex:1;}.gap-list .content strong{display:block;margin-bottom:2px;}.gap-list .content span{font-size:13px;color:var(--text-dim);}
.chart-box{background:var(--card-light);border:1px solid var(--border);border-radius:10px;padding:20px;margin:16px 0;}
.footer{text-align:center;padding:40px 24px;border-top:1px solid var(--border);margin-top:48px;font-size:13px;color:var(--text-mute);}
.footer p{margin-bottom:4px;}
@media print{body{background:#fff;color:#000;}.report-header{background:#f5f5f5;}.card,.method-card{background:#fafafa;border-color:#ddd;}.kpi-card{background:#fafafa;}}
</style>"""

    # 数据源更新时间
    try:
        mtime = os.path.getmtime(SRC)
        src_stamp = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
    except Exception:
        src_stamp = "未知"

    HTML = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>{C['report_title']}</title>{CSS}</head><body>
<div class="report-header">
  <div class="badge">生产车间 · 产能基准</div>
  <h1>{C['report_title']}</h1>
  <p class="subtitle">基于生产订单数据｜为车间负责人薪资提成设计提供产能基准测算</p>
  <div class="meta">
    <span>📊 数据来源：{os.path.basename(SRC)}</span>
    <span>📅 生成日期：{C['report_gen_date']}</span>
    <span>🏭 结算周期：按月</span>
    <span>⚙️ 完整运营月：5 个（3–7月）</span>
  </div>
</div>
<div class="container">

<div class="section">
  <div class="section-title"><span class="num">0</span> 执行摘要</div>
  <div class="section-desc">先说结论，再展开细节</div>
  <div class="kpi-grid">{kpi}</div>
  <div class="alert alert-warning"><strong>⚠️ 核心结论</strong>
  平衡点用于车间负责人提成设计——实际月产量超过 B 的部分对应提成。本报告<b>不替管理层拍板固定 B</b>，而是给出多算法候选值与敏感性对比，由管理层在候选中选定。</div>
  <div class="alert alert-info"><strong>📌 初步测算区间</strong>
  完整月数据下，平衡点候选区间为 <b>{fmt(raw_med)} – {fmt(norm_med)} 个/月</b>：
  取"历史实际达成"用<b>原始中位数 {fmt(raw_med)}</b>；取"纯产能"用<b>工作日标准化中位数 {fmt(norm_med)}</b>。两者均可通过后续滚动机制动态调整。</div>
</div>

<div class="section">
  <div class="section-title"><span class="num">1</span> 数据现状盘点</div>
  <div class="section-desc">手里有什么数据，哪些月份能用于建模</div>
  <div class="card">
    <div class="card-title"><span class="icon">📅</span> 月度产量与完整度</div>
    <table><thead><tr><th>月份</th><th class="text-right">产量（个）</th><th class="text-right">记录天数</th><th>日期跨度</th><th>完整度</th></tr></thead>
    <tbody>{month_rows}</tbody></table>
    <div class="alert alert-info" style="margin-top:12px;"><strong>📌 建模口径</strong>1/2/8 月天数缺失（春节、月初截数），建模仅用 <b>3–7 月</b> 五个完整运营月。注意各月<b>记录天数不同（21–27天）</b>——这是"工作日标准化法"的动机。</div>
  </div>
</div>

<div class="section">
  <div class="section-title"><span class="num">2</span> 月度产量分布</div>
  <div class="section-desc">完整月实际产量 vs 两种平衡点基准</div>
  <div class="chart-box">{svg}
    <div class="legend" style="display:flex;gap:20px;justify-content:center;margin-top:14px;font-size:12px;color:var(--text-dim);">
      <div style="display:flex;align-items:center;gap:6px;"><div style="width:10px;height:10px;border-radius:2px;background:#e85d3a;"></div>实际月产量</div>
      <div style="display:flex;align-items:center;gap:6px;"><div style="width:14px;height:0;border-top:2px dashed #3ecf8e;"></div>原始中位 B</div>
      <div style="display:flex;align-items:center;gap:6px;"><div style="width:14px;height:0;border-top:2px dashed #f0a030;"></div>标准化中位 B</div>
    </div>
  </div>
</div>

<div class="section">
  <div class="section-title"><span class="num">3</span> 平衡点 B 测算方法</div>
  <div class="section-desc">四种测算思路，交叉验证确定最终平衡点</div>
  {methods}
</div>

<div class="section">
  <div class="section-title"><span class="num">4</span> B 候选值与敏感性对比</div>
  <div class="section-desc">每个候选 B 下"几个月会超产、累计超产多少"，反映基线松紧</div>
  <div class="card">
    <div class="card-title"><span class="icon">🎯</span> 候选 B 取值</div>
    <table><thead><tr><th>算法</th><th class="text-right">候选 B（个/月）</th><th>说明</th></tr></thead><tbody>{cand_tbl}</tbody></table>
  </div>
  <div class="card">
    <div class="card-title"><span class="icon">📊</span> 敏感性：基线松紧</div>
    <table><thead><tr><th>候选 B</th><th class="text-right">取值</th><th class="text-right">触发超产月数</th><th class="text-right">累计超产总量（个）</th></tr></thead><tbody>{sens_tbl}</tbody></table>
    <div class="alert alert-success" style="margin-top:12px;"><strong>✅ 读法</strong>B 越低越容易触发提成（松），越高越难（紧）。若选原始中位 <b>{fmt(raw_med)}</b>，5 个月里 2 个月超产、累计超 {fmt(sens[1][3])} 个；若选标准化中位 <b>{fmt(norm_med)}</b>，仅 1 个月超产、累计 {fmt(sens[5][3])} 个。</div>
  </div>
</div>

<div class="section">
  <div class="section-title"><span class="num">5</span> 动态调整机制</div>
  <div class="section-desc">B 可能随时间变化，用滚动基线避免每次重拍</div>
  <div class="card">{timeline}</div>
</div>

<div class="section">
  <div class="section-title"><span class="num">6</span> 数据质量问题</div>
  <div class="section-desc">正式跑数 / 结算前须修复，否则污染基线</div>
  <div class="alert alert-danger"><strong>🚨 数据质量风险（3 项）</strong></div>
  <ul class="gap-list">{issues}</ul>
</div>

<div class="section">
  <div class="section-title"><span class="num">7</span> 决策建议</div>
  <div class="section-desc">不替你定，给决策路径</div>
  <div class="alert alert-success"><strong>✅ 建议决策路径</strong>
  <ol style="margin-left:20px;margin-top:8px;font-size:13px;line-height:1.9;">
    <li><b>选口径</b>：要"实际达成"选原始中位 {fmt(raw_med)}；要"纯产能"选标准化中位 {fmt(norm_med)}。</li>
    <li><b>定松紧</b>：依提成预算，在第四节敏感性表中微调（偏松选保守/比例法下界，偏紧选标准化法）。</li>
    <li><b>启用滚动</b>：选定初值后启用第五节滚动机制，后续自动动态可调。</li>
    <li><b>补规则</b>：提成分档（超产率阈值 + 单件提成）待管理层定；规则到位后本报告可扩展为提成测算表。</li>
  </ol></div>
</div>

</div>
<div class="footer">
  <p>本报告基于 {os.path.basename(SRC)} 数据生成（源文件更新于 {src_stamp}）</p>
  <p>分析框架设计：WorkBuddy ｜ 生成日期：{C['report_gen_date']}</p>
  <p>注：报告中的测算值基于有限样本（5 个完整运营月），供管理层框架讨论与决策使用</p>
</div>
</body></html>"""

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(HTML)
    print("OK ->", OUT)
    print("原始中位", raw_med, "标准化中位", norm_med, "总完整产量", total_cm)


if __name__ == "__main__":
    main()
