"""
store_health_html.py ─ Generate interactive HTML dashboard from Store Health analysis results.
Designed for publishing to GitHub Pages.
"""

import os, json
import pandas as pd
import numpy as np
from datetime import datetime


def generate_html_report(results, output_path):
    """Generate a self-contained HTML dashboard from analysis results."""
    print(f"\n[HTML] Generating HTML report...")

    df_frames = results['df_frames']
    df_contact = results['df_contact']
    brand_df = results['brand_df']
    retail_stores = results['retail_stores']
    summary = results['summary']
    sales_days = summary['sales_days']
    cl_sales_days = summary['cl_sales_days']
    n_stores = len(retail_stores)
    now = datetime.now().strftime('%Y-%m-%d %H:%M')

    # ── Prepare data for HTML ──────────────────────────────────────────

    # 1. Per-store Top N shortage %
    store_shortage = []
    for store in sorted(retail_stores):
        sf = df_frames[df_frames['store_name'] == store]
        sf_sorted = sf.sort_values('national_rank', ascending=True)
        row = {'store': store}
        for top_n in [50, 100, 200, 300]:
            tier = sf_sorted.head(top_n)
            n = len(tier)
            if n == 0:
                row[f'top{top_n}'] = None
            else:
                short = int(tier['is_stockout'].sum()) + int(tier['is_low_stock'].sum())
                row[f'top{top_n}'] = round(short / n * 100, 1)
        store_shortage.append(row)
    store_shortage.sort(key=lambda x: x.get('top50', 0) or 0, reverse=True)

    # National shortage
    nat_sorted = df_frames.sort_values('national_rank', ascending=True)
    nat_shortage = {'store': '全國'}
    for top_n in [50, 100, 200, 300]:
        tier = nat_sorted.head(top_n)
        n = len(tier)
        short = int(tier['is_stockout'].sum()) + int(tier['is_low_stock'].sum())
        nat_shortage[f'top{top_n}'] = round(short / n * 100, 1) if n > 0 else None

    # 2. Brand displayable matrix
    brand_matrix = []
    brand_list = []
    if not brand_df.empty:
        pivot = brand_df.pivot_table(
            index='store_name', columns='ブランド',
            values='displayable_sku_count', aggfunc='sum', fill_value=0)
        brand_list = pivot.sum().sort_values(ascending=False).index.tolist()
        pivot = pivot.reindex(columns=brand_list).fillna(0).astype(int)

        for store in [s['store'] for s in store_shortage]:
            row = {'store': store}
            for b in brand_list:
                row[b] = int(pivot.loc[store, b]) if store in pivot.index else 0
            brand_matrix.append(row)

    # 3. National top SKU list (top 300)
    nat_frames = df_frames.groupby('PLU', as_index=False).agg(
        品番=('品番', 'first'),
        ブランド=('ブランド', 'first') if 'ブランド' in df_frames.columns else ('品番', 'first'),
        sales=('sales', 'sum'),
        inventory=('inventory', 'sum'),
        national_rank=('national_rank', 'first'),
        oos_stores=('is_stockout', 'sum'),
        low_stores=('is_low_stock', 'sum'),
        n_stores=('store_name', 'nunique'),
    ).sort_values('national_rank').head(300)
    nat_frames['doh_m'] = np.where(
        nat_frames['sales'] > 0,
        (nat_frames['inventory'] / (nat_frames['sales'] / sales_days)) / 30, 999)

    nat_sku_data = []
    for _, r in nat_frames.iterrows():
        nat_sku_data.append({
            'rank': int(r['national_rank']),
            'hinban': r['品番'],
            'brand': r.get('ブランド', ''),
            'sales': int(r['sales']),
            'inv': int(r['inventory']),
            'doh': round(r['doh_m'], 1) if r['doh_m'] < 999 else '∞',
            'oos': int(r['oos_stores']),
            'low': int(r['low_stores']),
            'stores': int(r['n_stores']),
        })

    # ── Build HTML ─────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OWNDAYS TW Store Health Dashboard</title>
<style>
:root {{
    --blue: #1a73e8; --green: #0b8043; --red: #cc0000;
    --orange: #e67c00; --gray: #f8f9fa; --border: #dee2e6;
}}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; color: #333; }}
.container {{ max-width: 1400px; margin: 0 auto; padding: 20px; }}
header {{ background: linear-gradient(135deg, #1a73e8, #0b8043); color: #fff;
         padding: 24px 32px; border-radius: 12px; margin-bottom: 24px; }}
header h1 {{ font-size: 1.6em; margin-bottom: 4px; }}
header .meta {{ opacity: 0.85; font-size: 0.9em; }}
.kpi-row {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
           gap: 16px; margin-bottom: 24px; }}
.kpi {{ background: #fff; border-radius: 10px; padding: 18px; text-align: center;
       box-shadow: 0 1px 4px rgba(0,0,0,0.08); }}
.kpi .value {{ font-size: 2em; font-weight: 700; color: var(--blue); }}
.kpi .label {{ font-size: 0.85em; color: #666; margin-top: 4px; }}
.section {{ background: #fff; border-radius: 10px; padding: 24px;
           box-shadow: 0 1px 4px rgba(0,0,0,0.08); margin-bottom: 24px; }}
.section h2 {{ font-size: 1.2em; color: var(--blue); margin-bottom: 16px;
              border-bottom: 2px solid var(--blue); padding-bottom: 8px; }}
.section h2.green {{ color: var(--green); border-color: var(--green); }}
table {{ width: 100%; border-collapse: collapse; font-size: 0.85em; }}
th {{ background: var(--blue); color: #fff; padding: 8px 6px; text-align: center;
     position: sticky; top: 0; z-index: 2; }}
th.green {{ background: var(--green); }}
td {{ padding: 6px; text-align: center; border-bottom: 1px solid var(--border); }}
tr:hover {{ background: #f0f7ff; }}
.pct-high {{ background: #fce8e6; color: var(--red); font-weight: 700; }}
.pct-mid {{ background: #fef7e0; color: var(--orange); font-weight: 600; }}
.pct-low {{ background: #e6f4ea; color: var(--green); }}
.val-zero {{ color: var(--red); font-weight: 700; }}
.val-low {{ color: var(--orange); }}
.val-ok {{ color: var(--green); }}
.sticky-col {{ position: sticky; left: 0; background: #fff; z-index: 1;
              text-align: left !important; font-weight: 600; }}
tr:hover .sticky-col {{ background: #f0f7ff; }}
.nat-row {{ background: #e8f0fe !important; font-weight: 700; }}
.nat-row .sticky-col {{ background: #e8f0fe !important; }}
.scroll-wrapper {{ overflow-x: auto; max-height: 600px; overflow-y: auto; }}
.filter-bar {{ margin-bottom: 12px; display: flex; gap: 12px; align-items: center; }}
.filter-bar input {{ padding: 6px 12px; border: 1px solid var(--border); border-radius: 6px;
                    font-size: 0.9em; width: 220px; }}
.filter-bar select {{ padding: 6px 8px; border: 1px solid var(--border); border-radius: 6px; }}
.tab-bar {{ display: flex; gap: 4px; margin-bottom: 16px; flex-wrap: wrap; }}
.tab-bar button {{ padding: 8px 16px; border: 1px solid var(--border); background: #fff;
                  border-radius: 6px 6px 0 0; cursor: pointer; font-size: 0.9em; }}
.tab-bar button.active {{ background: var(--blue); color: #fff; border-color: var(--blue); }}
.tab-content {{ display: none; }}
.tab-content.active {{ display: block; }}
</style>
</head>
<body>
<div class="container">

<header>
  <h1>OWNDAYS Taiwan ─ Store Health Dashboard</h1>
  <div class="meta">更新時間: {now} ｜ 鏡框銷售期間: {sales_days}天 ｜ CL銷售期間: {cl_sales_days}天 ｜ 門市數: {n_stores}</div>
</header>

<div class="kpi-row">
  <div class="kpi"><div class="value">{n_stores}</div><div class="label">分析門市數</div></div>
  <div class="kpi"><div class="value">{summary['total_frame_skus']}</div><div class="label">鏡框 SKU 數</div></div>
  <div class="kpi"><div class="value">{summary['frame_stockout_pct']:.1f}%</div><div class="label">鏡框缺貨率 (qty≤1)</div></div>
  <div class="kpi"><div class="value">{summary['frame_low_stock_pct']:.1f}%</div><div class="label">庫存緊張率 (DOH&lt;1.5月)</div></div>
  <div class="kpi"><div class="value">{int(summary['dead_stock_count'])}</div><div class="label">無用庫存筆數</div></div>
  <div class="kpi"><div class="value">{summary['total_cl_skus']}</div><div class="label">CL SKU 數</div></div>
</div>

<div class="tab-bar">
  <button class="active" onclick="showTab('tab-shortage')">各門市庫存不足%</button>
  <button onclick="showTab('tab-brand')">Brand 可展示SKU</button>
  <button onclick="showTab('tab-national')">全國 Top SKU</button>
</div>

<!-- Tab 1: Shortage -->
<div id="tab-shortage" class="tab-content active">
<div class="section">
  <h2>各門市庫存不足% (缺貨 + 庫存緊張)</h2>
  <div class="filter-bar">
    <input type="text" id="filter-store" placeholder="搜尋門市..." oninput="filterTable('shortage-table', 1, this.value)">
  </div>
  <div class="scroll-wrapper">
  <table id="shortage-table">
  <thead><tr><th>No.</th><th>門市</th><th>Top 50</th><th>Top 100</th><th>Top 200</th><th>Top 300</th></tr></thead>
  <tbody>
  <tr class="nat-row"><td>—</td><td class="sticky-col">全國</td>"""

    for k in ['top50', 'top100', 'top200', 'top300']:
        v = nat_shortage[k]
        cls = 'pct-high' if v and v > 60 else ('pct-mid' if v and v > 40 else 'pct-low')
        html += f'<td class="{cls}">{v:.1f}%</td>' if v is not None else '<td>—</td>'
    html += '</tr>\n'

    for i, s in enumerate(store_shortage, 1):
        html += f'<tr><td>{i}</td><td class="sticky-col">{s["store"]}</td>'
        for k in ['top50', 'top100', 'top200', 'top300']:
            v = s[k]
            if v is not None:
                cls = 'pct-high' if v > 60 else ('pct-mid' if v > 40 else 'pct-low')
                html += f'<td class="{cls}">{v:.1f}%</td>'
            else:
                html += '<td>—</td>'
        html += '</tr>\n'

    html += """</tbody></table></div></div></div>

<!-- Tab 2: Brand Matrix -->
<div id="tab-brand" class="tab-content">
<div class="section">
  <h2 class="green">各門市 Brand 可展示 SKU 數 (在庫≥2)</h2>
  <div class="filter-bar">
    <input type="text" id="filter-brand-store" placeholder="搜尋門市..." oninput="filterTable('brand-table', 1, this.value)">
  </div>
  <div class="scroll-wrapper">
  <table id="brand-table">
  <thead><tr><th>No.</th><th>門市</th>"""

    for b in brand_list:
        html += f'<th class="green">{b}</th>'
    html += '</tr></thead>\n<tbody>\n'

    for i, row in enumerate(brand_matrix, 1):
        html += f'<tr><td>{i}</td><td class="sticky-col">{row["store"]}</td>'
        for b in brand_list:
            v = row.get(b, 0)
            cls = 'val-zero' if v == 0 else ('val-low' if v <= 2 else 'val-ok')
            html += f'<td class="{cls}">{v}</td>'
        html += '</tr>\n'

    # Total row
    html += '<tr class="nat-row"><td>—</td><td class="sticky-col">全國合計</td>'
    for b in brand_list:
        total = sum(row.get(b, 0) for row in brand_matrix)
        html += f'<td>{total}</td>'
    html += '</tr>\n'

    html += """</tbody></table></div></div></div>

<!-- Tab 3: National Top SKU -->
<div id="tab-national" class="tab-content">
<div class="section">
  <h2>全國鏡框 Top SKU 清單 (按銷量排名)</h2>
  <div class="filter-bar">
    <input type="text" id="filter-sku" placeholder="搜尋品番 / 品牌..." oninput="filterTable('national-table', 1, this.value, [1,2])">
    <select onchange="filterNatStatus(this.value)">
      <option value="">全部</option>
      <option value="oos">缺貨店數 > 0</option>
      <option value="low">庫存緊張 > 0</option>
    </select>
  </div>
  <div class="scroll-wrapper" style="max-height:700px;">
  <table id="national-table">
  <thead><tr>
    <th>排名</th><th>品番</th><th>品牌</th><th>全台銷量</th><th>全台庫存</th>
    <th>DOH(月)</th><th>缺貨店數</th><th>緊張店數</th><th>鋪貨店數</th>
  </tr></thead>
  <tbody>"""

    for sku in nat_sku_data:
        oos_cls = ' class="pct-high"' if sku['oos'] > n_stores * 0.3 else (' class="pct-mid"' if sku['oos'] > 0 else '')
        html += f"""<tr data-oos="{sku['oos']}" data-low="{sku['low']}">
  <td>{sku['rank']}</td><td style="text-align:left">{sku['hinban']}</td>
  <td>{sku['brand']}</td><td>{sku['sales']}</td><td>{sku['inv']}</td>
  <td>{sku['doh']}</td><td{oos_cls}>{sku['oos']}</td>
  <td>{sku['low']}</td><td>{sku['stores']}</td></tr>\n"""

    html += """</tbody></table></div></div></div>

</div><!-- container -->

<script>
function showTab(id) {
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-bar button').forEach(b => b.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  event.target.classList.add('active');
}

function filterTable(tableId, colIdx, query, cols) {
  const rows = document.querySelectorAll('#' + tableId + ' tbody tr');
  const q = query.toLowerCase();
  rows.forEach(row => {
    if (row.classList.contains('nat-row')) return;
    const searchCols = cols || [colIdx];
    let match = false;
    searchCols.forEach(ci => {
      const cell = row.cells[ci];
      if (cell && cell.textContent.toLowerCase().includes(q)) match = true;
    });
    row.style.display = match ? '' : 'none';
  });
}

function filterNatStatus(val) {
  const rows = document.querySelectorAll('#national-table tbody tr');
  rows.forEach(row => {
    if (!val) { row.style.display = ''; return; }
    const oos = parseInt(row.dataset.oos || '0');
    const low = parseInt(row.dataset.low || '0');
    if (val === 'oos') row.style.display = oos > 0 ? '' : 'none';
    else if (val === 'low') row.style.display = low > 0 ? '' : 'none';
  });
}
</script>
</body>
</html>"""

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"    -> HTML report saved: {output_path}")
    return output_path
