import json
import math
import re
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "各省光伏现货电价数据.xlsx"
OUTPUT_DIR = ROOT / "outputs"
OUTPUT = OUTPUT_DIR / "各省光伏现货电价数据看板.html"
REFERENCE_URL = "https://mp.weixin.qq.com/s/2UPt0Y0rkwJH3HBpTzNDkA"


def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m")
    text = str(value).strip()
    if text == "nan":
        return ""
    return text


def first_number(value):
    text = clean_text(value)
    if not text:
        return None
    for match in re.finditer(r"-?\d+(?:\.\d+)?", text):
        number = float(match.group(0))
        if "%" in text[match.start() : match.end() + 1] and abs(number) > 1:
            number = number / 100
        # Electricity prices and ratios in this workbook are all well below 1.5.
        # This skips years/month markers such as 2026年 or 25年 in prose cells.
        if abs(number) <= 1.5:
            return number
    return None


def latest_numeric(*values):
    for value in reversed(values):
        number = first_number(value)
        if number is not None:
            return number
    return None


def load_records():
    raw = pd.read_excel(INPUT, sheet_name=0, header=None)
    records = []
    region = ""
    for _, row in raw.iloc[3:].iterrows():
        if clean_text(row[1]):
            region = clean_text(row[1])
        province = clean_text(row[2])
        if not province:
            continue

        monthly = {
            "2026-01": first_number(row[17]),
            "2026-02": first_number(row[18]),
            "2026-03": first_number(row[19]),
            "2026-04": first_number(row[20]),
            "2026-05": first_number(row[21]),
        }
        spot_windows = {
            "2025": first_number(row[10]),
            "2025.4-2026.3": first_number(row[11]),
            "2025.5-2026.4": first_number(row[12]),
            "2025.6-2026.5": first_number(row[13]),
        }
        record = {
            "region": region or "未分区",
            "province": province,
            "coal": first_number(row[3]),
            "coalRaw": clean_text(row[3]),
            "stockRatio": clean_text(row[4]),
            "flow": clean_text(row[5]),
            "bid2025": first_number(row[6]),
            "bid2025Raw": clean_text(row[6]),
            "bid2025Ratio": clean_text(row[7]),
            "bid2026": first_number(row[8]),
            "bid2026Raw": clean_text(row[8]),
            "bid2026Ratio": clean_text(row[9]),
            "spot2025Raw": clean_text(row[10]),
            "spotLatest": latest_numeric(row[10], row[11], row[12], row[13]),
            "spotWindows": spot_windows,
            "income": first_number(row[14]),
            "incomeRaw": clean_text(row[14]),
            "node": clean_text(row[15]) or "未说明",
            "allocation": clean_text(row[16]),
            "monthly": monthly,
            "settlementLatest": latest_numeric(row[17], row[18], row[19], row[20], row[21]),
            "settlementRaw": {
                "2026-01": clean_text(row[17]),
                "2026-02": clean_text(row[18]),
                "2026-03": clean_text(row[19]),
                "2026-04": clean_text(row[20]),
                "2026-05": clean_text(row[21]),
            },
            "acNote": clean_text(row[22]),
            "source": clean_text(row[23]),
            "remark": clean_text(row[24]),
        }
        records.append(record)
    return records


def compact_json(data):
    return json.dumps(data, ensure_ascii=False, separators=(",", ":"))


def build_html(records):
    data_json = compact_json(records)
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>各省光伏现货电价数据看板</title>
<style>
:root {{
  color-scheme: light;
  --ink:#18212f; --muted:#637083; --line:#d9e0e8; --paper:#f6f8fb;
  --panel:#ffffff; --green:#087f5b; --blue:#2563eb; --gold:#b7791f; --red:#c2410c;
  --cyan:#0891b2; --soft-green:#e8f6ef; --soft-blue:#eaf1ff; --soft-gold:#fff4dc; --soft-red:#fff0e8;
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",Arial,sans-serif;
}}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--paper); color:var(--ink); }}
header {{ position:relative; padding:24px 18px 18px; background:linear-gradient(135deg,#0d5d56 0%,#1c6aa6 58%,#604c95 100%); color:#fff; overflow:hidden; }}
header:after {{ content:""; position:absolute; inset:auto -18% -45% 32%; height:160px; background:rgba(255,255,255,.14); transform:rotate(-8deg); }}
.hero {{ position:relative; z-index:1; max-width:1180px; margin:0 auto; }}
.eyebrow {{ font-size:12px; opacity:.82; letter-spacing:0; }}
h1 {{ margin:8px 0 8px; font-size:clamp(26px,6vw,48px); line-height:1.08; letter-spacing:0; }}
.subtitle {{ max-width:760px; margin:0; font-size:14px; line-height:1.65; opacity:.9; }}
.wrap {{ max-width:1180px; margin:0 auto; padding:14px; }}
.filters {{ position:sticky; top:0; z-index:3; display:grid; grid-template-columns:1.5fr repeat(3,1fr); gap:8px; padding:10px 14px; background:rgba(246,248,251,.92); backdrop-filter:blur(10px); border-bottom:1px solid var(--line); }}
input,select,button {{ width:100%; min-height:40px; border:1px solid var(--line); border-radius:8px; background:#fff; color:var(--ink); font:inherit; padding:0 10px; }}
button {{ cursor:pointer; font-weight:650; }}
.kpis {{ display:grid; grid-template-columns:repeat(4,1fr); gap:10px; margin:12px 0; }}
.kpi,.panel,.province {{ background:var(--panel); border:1px solid var(--line); border-radius:8px; box-shadow:0 8px 24px rgba(15,30,50,.06); }}
.kpi {{ padding:13px; min-height:102px; }}
.kpi span {{ display:block; color:var(--muted); font-size:12px; }}
.kpi strong {{ display:block; margin-top:8px; font-size:28px; line-height:1; }}
.kpi em {{ display:block; margin-top:8px; color:var(--muted); font-style:normal; font-size:12px; line-height:1.35; }}
.grid {{ display:grid; grid-template-columns:1.08fr .92fr; gap:12px; }}
.panel {{ padding:14px; overflow:hidden; }}
.panel h2 {{ margin:0 0 10px; font-size:16px; line-height:1.25; }}
.hint {{ color:var(--muted); font-size:12px; margin-top:-4px; margin-bottom:10px; }}
.bars {{ display:grid; gap:8px; }}
.bar {{ display:grid; grid-template-columns:88px 1fr 54px; align-items:center; gap:8px; font-size:12px; }}
.bar .name {{ overflow:hidden; white-space:nowrap; text-overflow:ellipsis; }}
.track {{ height:12px; background:#edf2f7; border-radius:8px; overflow:hidden; }}
.fill {{ height:100%; border-radius:8px; background:linear-gradient(90deg,var(--green),var(--cyan)); }}
.value {{ text-align:right; color:var(--muted); font-variant-numeric:tabular-nums; }}
.scatter,.linechart {{ width:100%; height:260px; display:block; background:#fff; }}
.legend {{ display:flex; flex-wrap:wrap; gap:8px; color:var(--muted); font-size:12px; }}
.dot {{ width:9px; height:9px; display:inline-block; border-radius:50%; margin-right:4px; vertical-align:-1px; }}
.provinces {{ display:grid; grid-template-columns:repeat(2,1fr); gap:10px; margin-top:12px; }}
.province {{ padding:12px; display:grid; gap:9px; }}
.phead {{ display:flex; justify-content:space-between; gap:10px; align-items:flex-start; }}
.phead strong {{ font-size:16px; }}
.tagrow {{ display:flex; flex-wrap:wrap; gap:6px; }}
.tag {{ border-radius:999px; padding:4px 8px; font-size:11px; color:#314156; background:#edf2f7; }}
.tag.node {{ background:var(--soft-blue); color:#1d4ed8; }}
.tag.recv {{ background:var(--soft-green); color:#047857; }}
.tag.send {{ background:var(--soft-gold); color:#9a5b00; }}
.metrics {{ display:grid; grid-template-columns:repeat(3,1fr); gap:6px; }}
.metric {{ background:#f8fafc; border:1px solid #edf2f7; border-radius:8px; padding:8px; min-width:0; }}
.metric span {{ display:block; color:var(--muted); font-size:11px; }}
.metric b {{ display:block; margin-top:4px; font-size:15px; font-variant-numeric:tabular-nums; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
.note {{ color:#415066; font-size:12px; line-height:1.55; display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; overflow:hidden; }}
.more {{ border-color:#cbd5e1; color:#245276; background:#fff; }}
.drawer {{ position:fixed; inset:auto 0 0; z-index:9; background:#fff; border-radius:16px 16px 0 0; box-shadow:0 -18px 50px rgba(0,0,0,.22); padding:18px; transform:translateY(110%); transition:.22s ease; max-height:84vh; overflow:auto; }}
.drawer.show {{ transform:translateY(0); }}
.shade {{ position:fixed; inset:0; background:rgba(15,23,42,.34); z-index:8; display:none; }}
.shade.show {{ display:block; }}
.detail-grid {{ display:grid; grid-template-columns:repeat(2,1fr); gap:8px; margin:12px 0; }}
.detail-grid div {{ background:#f8fafc; border:1px solid #edf2f7; border-radius:8px; padding:9px; font-size:13px; }}
.detail-grid span {{ display:block; color:var(--muted); font-size:11px; margin-bottom:4px; }}
.source {{ overflow-wrap:anywhere; color:#2563eb; font-size:12px; }}
footer {{ color:var(--muted); font-size:12px; line-height:1.6; padding:18px 14px 28px; max-width:1180px; margin:0 auto; }}
@media (max-width:820px) {{
  .filters {{ grid-template-columns:1fr 1fr; }}
  .filters input {{ grid-column:1 / -1; }}
  .kpis,.grid,.provinces {{ grid-template-columns:1fr; }}
  .kpi strong {{ font-size:24px; }}
  .panel {{ padding:12px; }}
  .bar {{ grid-template-columns:74px 1fr 48px; }}
  .scatter,.linechart {{ height:220px; }}
}}
</style>
</head>
<body>
<header><div class="hero">
  <div class="eyebrow">数据源：各省光伏现货电价数据.xlsx · 更新口径见原表备注</div>
  <h1>各省光伏现货电价数据看板</h1>
  <p class="subtitle">把燃煤基准、机制竞价、现货价格、节点电价与结算单参考价放在同一屏里比较。适合手机浏览和微信文件转发，所有图表均为本地渲染，无需联网。</p>
</div></header>
<section class="filters">
  <input id="search" placeholder="搜索省份 / 备注 / A-C 说明">
  <select id="region"><option value="">全部区域</option></select>
  <select id="node"><option value="">节点电价：全部</option></select>
  <select id="flow"><option value="">送受电：全部</option></select>
</section>
<main class="wrap">
  <section class="kpis" id="kpis"></section>
  <section class="grid">
    <div class="panel">
      <h2>近期期现货价格低位省份</h2>
      <div class="hint">按“2025.6-2026.5现货价格”优先，缺失时回看前序窗口。</div>
      <div class="bars" id="rankBars"></div>
    </div>
    <div class="panel">
      <h2>燃煤基准 vs 现货参考</h2>
      <div class="hint">横轴为燃煤基准电价，纵轴为最新可用现货价格。</div>
      <svg class="scatter" id="scatter" role="img"></svg>
    </div>
    <div class="panel">
      <h2>区域现货均值</h2>
      <div class="hint">只纳入可解析为数字的省份。</div>
      <div class="bars" id="regionBars"></div>
    </div>
    <div class="panel">
      <h2>结算单月度参考价</h2>
      <div class="hint">点击下方任一省份卡片，可切换这里的月度折线。</div>
      <svg class="linechart" id="linechart" role="img"></svg>
      <div class="legend" id="lineLegend"></div>
    </div>
  </section>
  <section class="provinces" id="cards"></section>
</main>
<footer>
  <div>参考公众号链接：{REFERENCE_URL}</div>
  <div>说明：数值由原表文本中首个数字解析而来，类似“无现货、未获取到数据、无结算单”的状态保留在省份详情中；排行榜和均值只统计可解析数字。</div>
</footer>
<div class="shade" id="shade"></div>
<aside class="drawer" id="drawer"></aside>
<script>
const DATA = {data_json};
const fmt = v => v==null ? "—" : Number(v).toFixed(4);
const pct = v => v==null ? "—" : (Number(v)*100).toFixed(0)+"%";
const avg = arr => arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : null;
const uniq = arr => [...new Set(arr.filter(Boolean))];
const colorByRegion = r => {{
  const map = {{华北:"#087f5b",东北:"#2563eb",华东:"#c2410c",华中:"#7c3aed",华南:"#0891b2",西南:"#b7791f",西北:"#475569"}};
  return map[r] || "#64748b";
}};
let selected = DATA.find(d => d.monthly && Object.values(d.monthly).some(v => v!=null)) || DATA[0];

function initFilters() {{
  for (const r of uniq(DATA.map(d=>d.region))) region.add(new Option(r,r));
  for (const n of uniq(DATA.map(d=>d.node))) node.add(new Option(n,n));
  for (const f of uniq(DATA.map(d=>d.flow))) flow.add(new Option(f,f));
  [search,region,node,flow].forEach(el => el.addEventListener("input", render));
}}
function filtered() {{
  const q = search.value.trim().toLowerCase();
  return DATA.filter(d => (!region.value || d.region===region.value)
    && (!node.value || d.node===node.value)
    && (!flow.value || d.flow===flow.value)
    && (!q || [d.province,d.region,d.acNote,d.remark,d.allocation].join(" ").toLowerCase().includes(q)));
}}
function renderKpis(rows) {{
  const spots = rows.map(d=>d.spotLatest).filter(v=>v!=null);
  const nodes = rows.filter(d => /是|节点|分时/.test(d.node) && !/否/.test(d.node)).length;
  const incomes = rows.map(d=>d.income).filter(v=>v!=null);
  const low = spots.filter(v=>v < 0.18).length;
  const cards = [
    ["覆盖省份/电网", rows.length, "当前筛选后的可见条目"],
    ["现货均值", fmt(avg(spots)), `${{spots.length}} 个条目有可计算现货价`],
    ["低于 0.18 元/kWh", low, "按最新可用现货窗口统计"],
    ["节点/分时相关", nodes, "含“是、节点、分时”等表述"]
  ];
  kpis.innerHTML = cards.map((c,i)=>`<div class="kpi"><span>${{c[0]}}</span><strong>${{c[1]}}</strong><em>${{c[2]}}</em></div>`).join("");
}}
function renderBars(el, rows, valueKey, limit=10, asc=true) {{
  const items = rows.filter(d=>d[valueKey]!=null).sort((a,b)=>asc ? a[valueKey]-b[valueKey] : b[valueKey]-a[valueKey]).slice(0,limit);
  const max = Math.max(...items.map(d=>d[valueKey]), .01);
  el.innerHTML = items.length ? items.map(d=>`<div class="bar"><div class="name" title="${{d.province}}">${{d.province}}</div><div class="track"><div class="fill" style="width:${{Math.max(4,d[valueKey]/max*100)}}%;background:${{colorByRegion(d.region)}}"></div></div><div class="value">${{fmt(d[valueKey])}}</div></div>`).join("") : `<div class="hint">暂无可计算数据</div>`;
}}
function renderRegionBars(rows) {{
  const groups = new Map();
  rows.forEach(d => {{ if (d.spotLatest!=null) {{ if(!groups.has(d.region)) groups.set(d.region, []); groups.get(d.region).push(d.spotLatest); }} }});
  const items = [...groups.entries()].map(([region,vals]) => ({{province:region, region, spotLatest:avg(vals)}})).sort((a,b)=>a.spotLatest-b.spotLatest);
  renderBars(regionBars, items, "spotLatest", 8, true);
}}
function drawScatter(rows) {{
  const pts = rows.filter(d=>d.coal!=null && d.spotLatest!=null);
  const svg = scatter, w = svg.clientWidth || 520, h = svg.clientHeight || 260, p = 34;
  svg.setAttribute("viewBox", `0 0 ${{w}} ${{h}}`);
  if (!pts.length) {{ svg.innerHTML = ""; return; }}
  const xs = pts.map(d=>d.coal), ys = pts.map(d=>d.spotLatest);
  const minX=Math.min(...xs)*.94, maxX=Math.max(...xs)*1.04, minY=Math.min(...ys)*.88, maxY=Math.max(...ys)*1.08;
  const sx=x=>p+(x-minX)/(maxX-minX)*(w-p*1.5), sy=y=>h-p-(y-minY)/(maxY-minY)*(h-p*1.6);
  svg.innerHTML = `<line x1="${{p}}" y1="${{h-p}}" x2="${{w-p/2}}" y2="${{h-p}}" stroke="#cbd5e1"/><line x1="${{p}}" y1="${{p/2}}" x2="${{p}}" y2="${{h-p}}" stroke="#cbd5e1"/>`
    + pts.map(d=>`<circle cx="${{sx(d.coal)}}" cy="${{sy(d.spotLatest)}}" r="5" fill="${{colorByRegion(d.region)}}"><title>${{d.province}} 燃煤 ${{fmt(d.coal)}} / 现货 ${{fmt(d.spotLatest)}}</title></circle>`).join("")
    + `<text x="${{p}}" y="${{18}}" fill="#637083" font-size="11">现货</text><text x="${{w-70}}" y="${{h-8}}" fill="#637083" font-size="11">燃煤基准</text>`;
}}
function drawLine(d) {{
  const vals = Object.entries(d.monthly || {{}}).filter(([k,v])=>v!=null);
  const svg = linechart, w = svg.clientWidth || 520, h = svg.clientHeight || 260, p = 34;
  svg.setAttribute("viewBox", `0 0 ${{w}} ${{h}}`);
  lineLegend.innerHTML = `<span><span class="dot" style="background:${{colorByRegion(d.region)}}"></span>${{d.province}}</span><span>最新结算参考：${{fmt(d.settlementLatest)}}</span>`;
  if (vals.length < 2) {{ svg.innerHTML = `<text x="16" y="42" fill="#637083" font-size="13">该条目月度结算单数据不足</text>`; return; }}
  const ys = vals.map(x=>x[1]), minY=Math.min(...ys)*.82, maxY=Math.max(...ys)*1.12;
  const sx=i=>p+i/(vals.length-1)*(w-p*1.6), sy=y=>h-p-(y-minY)/(maxY-minY)*(h-p*1.7);
  const path = vals.map(([k,v],i)=>`${{i?"L":"M"}} ${{sx(i)}} ${{sy(v)}}`).join(" ");
  svg.innerHTML = `<line x1="${{p}}" y1="${{h-p}}" x2="${{w-p/2}}" y2="${{h-p}}" stroke="#cbd5e1"/><line x1="${{p}}" y1="${{p/2}}" x2="${{p}}" y2="${{h-p}}" stroke="#cbd5e1"/><path d="${{path}}" fill="none" stroke="${{colorByRegion(d.region)}}" stroke-width="3"/>`
    + vals.map(([k,v],i)=>`<g><circle cx="${{sx(i)}}" cy="${{sy(v)}}" r="4" fill="${{colorByRegion(d.region)}}"/><text x="${{sx(i)-18}}" y="${{h-10}}" fill="#637083" font-size="10">${{k.slice(5)}}</text><text x="${{sx(i)-18}}" y="${{sy(v)-8}}" fill="#334155" font-size="10">${{fmt(v)}}</text></g>`).join("");
}}
function renderCards(rows) {{
  cards.innerHTML = rows.map((d,i)=>`<article class="province">
    <div class="phead"><strong>${{d.province}}</strong><span class="tag">${{d.region}}</span></div>
    <div class="tagrow"><span class="tag ${{d.flow==="受入"?"recv":d.flow==="送出"?"send":""}}">${{d.flow||"送受电未填"}}</span><span class="tag node">${{d.node}}</span></div>
    <div class="metrics"><div class="metric"><span>燃煤基准</span><b>${{d.coalRaw||fmt(d.coal)}}</b></div><div class="metric"><span>机制竞价25</span><b>${{d.bid2025Raw||fmt(d.bid2025)}}</b></div><div class="metric"><span>近现货</span><b>${{fmt(d.spotLatest)}}</b></div></div>
    <div class="note">${{d.acNote || d.remark || "暂无备注"}}</div>
    <button class="more" onclick="openDetail('${{d.province.replace(/'/g,"\\'")}}')">查看详情</button>
  </article>`).join("");
}}
function openDetail(name) {{
  const d = DATA.find(x=>x.province===name);
  selected = d; drawLine(d);
  drawer.innerHTML = `<h2>${{d.province}} · ${{d.region}}</h2>
    <div class="tagrow"><span class="tag">${{d.flow||"送受电未填"}}</span><span class="tag node">${{d.node}}</span><span class="tag">存量比例 ${{d.stockRatio||"—"}}</span></div>
    <div class="detail-grid">
      <div><span>燃煤基准</span>${{d.coalRaw||fmt(d.coal)}}</div><div><span>2025机制竞价</span>${{d.bid2025Raw||"—"}} / ${{d.bid2025Ratio||"—"}}</div>
      <div><span>2026机制竞价</span>${{d.bid2026Raw||"—"}} / ${{d.bid2026Ratio||"—"}}</div><div><span>聚合增收预期</span>${{d.incomeRaw||fmt(d.income)}}</div>
      <div><span>最新现货参考</span>${{fmt(d.spotLatest)}}</div><div><span>最新结算单参考</span>${{fmt(d.settlementLatest)}}</div>
    </div>
    <p><b>分摊或补偿项</b><br>${{d.allocation || "—"}}</p>
    <p><b>A 与 C 口径</b><br>${{d.acNote || "—"}}</p>
    <p><b>补充备注</b><br>${{d.remark || "—"}}</p>
    <p class="source"><b>来源链接</b><br>${{d.source || "原表未填"}}</p>
    <button onclick="closeDetail()">关闭</button>`;
  drawer.classList.add("show"); shade.classList.add("show");
}}
function closeDetail() {{ drawer.classList.remove("show"); shade.classList.remove("show"); }}
shade.addEventListener("click", closeDetail);
function render() {{
  const rows = filtered();
  renderKpis(rows); renderBars(rankBars, rows, "spotLatest", 10, true); renderRegionBars(rows); drawScatter(rows); renderCards(rows); drawLine(selected);
}}
window.addEventListener("resize", () => {{ drawScatter(filtered()); drawLine(selected); }});
initFilters(); render();
</script>
</body>
</html>"""


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    records = load_records()
    OUTPUT.write_text(build_html(records), encoding="utf-8")
    print(OUTPUT)
    print(f"records={len(records)}")


if __name__ == "__main__":
    main()
