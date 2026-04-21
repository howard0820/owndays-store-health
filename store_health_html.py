"""
store_health_html.py ─ Generate interactive HTML dashboard from Store Health analysis results.
Designed for publishing to GitHub Pages.
"""

import os, json
import pandas as pd
import numpy as np
from datetime import datetime, timedelta


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
    base_date_str = summary.get('base_date', '')
    n_stores = len(retail_stores)
    now = datetime.now().strftime('%Y-%m-%d %H:%M')

    # Date labels
    if base_date_str:
        bd = datetime.strptime(base_date_str, '%Y-%m-%d').date()
        frame_start = bd - timedelta(days=sales_days - 1)
        sales_date_label = f'{frame_start}~{bd}'
        inv_date_label = str(bd)
    else:
        sales_date_label = f'{sales_days}d'
        inv_date_label = ''

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
            row_total = 0
            for b in brand_list:
                v = int(pivot.loc[store, b]) if store in pivot.index else 0
                row[b] = v
                row_total += v
            row['_total'] = row_total
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

    # 4. Per-store data
    store_data = {}
    for store in sorted(retail_stores):
        sf = df_frames[df_frames['store_name'] == store].copy()
        sc = df_contact[df_contact['store_name'] == store].copy() if not df_contact.empty else pd.DataFrame()

        # Top N summary
        sf_by_rank = sf.sort_values('national_rank', ascending=True)
        tiers = {}
        for top_n in [50, 100, 200, 300]:
            tier = sf_by_rank.head(top_n)
            n = len(tier)
            if n == 0:
                tiers[top_n] = {'n': 0, 'oos': 0, 'low': 0, 'pct': None}
            else:
                oos = int(tier['is_stockout'].sum())
                low = int(tier['is_low_stock'].sum())
                tiers[top_n] = {'n': n, 'oos': oos, 'low': low, 'pct': round((oos+low)/n*100, 1)}

        # Frame SKU list (top 200 by store sales)
        sf_sorted = sf.sort_values('sales', ascending=False).head(200)
        frames = []
        for _, r in sf_sorted.iterrows():
            doh = round(r['doh_months'], 1) if r['doh_months'] < 9999 else '∞'
            if r['is_stockout']:
                status = '缺貨'
            elif r['is_low_stock']:
                status = '庫存緊張'
            else:
                status = '貨量OK'
            frames.append({
                'hinban': r['品番'], 'plu': r['PLU'],
                'brand': r.get('ブランド', ''),
                'sales': int(r['sales']), 'inv': int(r['inventory']),
                'doh': doh, 'nat_rank': int(r['national_rank']),
                'store_rank': int(r['store_rank']), 'status': status,
                'replenish': int(r['replenish_qty']) if r['needs_replenish'] else 0,
            })

        # CL list
        cls_list = []
        if not sc.empty:
            sc_sorted = sc.sort_values('sales', ascending=False).head(100)
            for _, r in sc_sorted.iterrows():
                doh = round(r['doh_months'], 1) if r['doh_months'] < 9999 else '∞'
                status = '缺貨' if r['is_stockout'] else ('庫存緊張' if r['is_low_stock'] else '貨量OK')
                cls_list.append({
                    'hinban': r['品番'], 'degree': r.get('degree', ''),
                    'plu': r['PLU'], 'sales': int(r['sales']),
                    'inv': int(r['inventory']), 'doh': doh, 'status': status,
                })

        # Brand display data
        brands = []
        if not brand_df.empty:
            sb = brand_df[brand_df['store_name'] == store].sort_values('total_sales', ascending=False)
            for _, r in sb.iterrows():
                doh = round(r['doh_months'], 1) if r['doh_months'] < 9999 else '∞'
                brands.append({
                    'brand': r['ブランド'],
                    'sku': int(r['sku_count']),
                    'display': int(r['displayable_sku_count']),
                    'inv': int(r['total_inv']),
                    'doh': doh,
                    'pct': round(r['sales_pct'] * 100, 1),
                })

        store_data[store] = {'tiers': tiers, 'frames': frames, 'cls': cls_list, 'brands': brands}

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
  <button onclick="showTab('tab-stores')" style="background:#FF6D01;color:#fff;border-color:#FF6D01">📦 各門市明細</button>
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
  <thead><tr><th>No.</th><th>門市</th><th class="green">Total</th>"""

    for b in brand_list:
        html += f'<th class="green">{b}</th>'
    html += '</tr></thead>\n<tbody>\n'

    for i, row in enumerate(brand_matrix, 1):
        html += f'<tr><td>{i}</td><td class="sticky-col">{row["store"]}</td>'
        html += f'<td><b>{row.get("_total", 0)}</b></td>'
        for b in brand_list:
            v = row.get(b, 0)
            cls = 'val-zero' if v == 0 else ('val-low' if v <= 2 else 'val-ok')
            html += f'<td class="{cls}">{v}</td>'
        html += '</tr>\n'

    # Total row
    grand_total = sum(row.get('_total', 0) for row in brand_matrix)
    html += f'<tr class="nat-row"><td>—</td><td class="sticky-col">全國合計</td><td><b>{grand_total}</b></td>'
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
    <th>排名</th><th>品番</th><th>品牌</th><th>全台銷量<br>(""" + sales_date_label + """)</th><th>全台庫存<br>(""" + inv_date_label + """)</th>
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

<!-- Tab 4: Per-store detail -->
<div id="tab-stores" class="tab-content">
<div class="section">
  <h2 style="border-color:#FF6D01;color:#FF6D01">📦 各門市明細</h2>
  <div class="filter-bar">
    <select id="store-select" onchange="showStore(this.value)" style="width:300px;font-size:1em;padding:8px;">
      <option value="">-- 選擇門市 --</option>"""

    # Store options sorted by Top50 shortage (worst first)
    for s in store_shortage:
        pct_label = f" ({s['top50']:.0f}%)" if s.get('top50') is not None else ''
        html += f'\n      <option value="{s["store"]}">{s["store"]}{pct_label}</option>'

    html += """
    </select>
  </div>
  <div id="store-detail-area" style="margin-top:16px;"></div>
</div>
</div>

</div><!-- container -->

<script>
const storeData = """

    # Serialize store data as JSON
    import json
    # Convert to JSON-safe format
    html += json.dumps(store_data, ensure_ascii=False)

    html += """;

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

function pctClass(v) { return v > 60 ? 'pct-high' : (v > 40 ? 'pct-mid' : 'pct-low'); }
function statusClass(s) { return s === '缺貨' ? 'pct-high' : (s === '庫存緊張' ? 'pct-mid' : ''); }

function showStore(name) {
  const area = document.getElementById('store-detail-area');
  if (!name || !storeData[name]) { area.innerHTML = ''; return; }
  const d = storeData[name];
  let h = '';

  // Summary
  h += '<h3 style="color:#FF6D01;margin-bottom:8px;">📊 ' + name + ' 庫存健康摘要</h3>';
  h += '<table style="width:auto;margin-bottom:20px"><thead><tr><th>排名區間</th><th>SKU數</th><th>缺貨數</th><th>庫存緊張數</th><th>庫存不足%</th></tr></thead><tbody>';
  [50,100,200,300].forEach(n => {
    const t = d.tiers[n];
    if (!t || t.n === 0) return;
    const cls = pctClass(t.pct);
    h += '<tr><td>Top ' + n + (t.n < n ? ' (實際'+t.n+')' : '') + '</td><td>' + t.n + '</td><td>' + t.oos + '</td><td>' + t.low + '</td><td class="' + cls + '">' + t.pct.toFixed(1) + '%</td></tr>';
  });
  h += '</tbody></table>';

  // Frame list
  h += '<h3 style="color:#1a73e8;margin-bottom:8px;">📦 鏡框 / 太陽眼鏡 SKU清單</h3>';
  h += '<div class="filter-bar"><input type="text" placeholder="搜尋品番/品牌..." oninput="filterStoreTable(this.value)"></div>';
  h += '<div class="scroll-wrapper" style="max-height:500px"><table id="store-frame-table"><thead><tr>';
  h += '<th>品番</th><th>PLU</th><th>品牌</th><th>銷量</th><th>在庫</th><th>DOH(月)</th><th>全台排名</th><th>店內排名</th><th>狀態</th><th>建議補貨</th>';
  h += '</tr></thead><tbody>';
  d.frames.forEach(f => {
    const sc = statusClass(f.status);
    const fill = f.status === '缺貨' ? 'style="background:#fce8e6"' : (f.status === '庫存緊張' ? 'style="background:#fef7e0"' : '');
    h += '<tr ' + fill + '><td style="text-align:left">' + f.hinban + '</td><td>' + f.plu + '</td><td>' + f.brand + '</td>';
    h += '<td>' + f.sales + '</td><td>' + f.inv + '</td><td>' + f.doh + '</td>';
    h += '<td>' + f.nat_rank + '</td><td>' + f.store_rank + '</td>';
    h += '<td class="' + sc + '">' + f.status + '</td>';
    h += '<td>' + (f.replenish > 0 ? f.replenish : '') + '</td></tr>';
  });
  h += '</tbody></table></div>';

  // CL list
  if (d.cls && d.cls.length > 0) {
    h += '<h3 style="color:#e67c00;margin:20px 0 8px;">🟡 隱形眼鏡 SKU清單</h3>';
    h += '<div class="scroll-wrapper" style="max-height:400px"><table><thead><tr>';
    h += '<th>品番</th><th>度數</th><th>PLU</th><th>銷量</th><th>在庫</th><th>DOH(月)</th><th>狀態</th>';
    h += '</tr></thead><tbody>';
    d.cls.forEach(c => {
      const sc = statusClass(c.status);
      const fill = c.status === '缺貨' ? 'style="background:#fce8e6"' : (c.status === '庫存緊張' ? 'style="background:#fef7e0"' : '');
      h += '<tr ' + fill + '><td style="text-align:left">' + c.hinban + '</td><td>' + c.degree + '</td><td>' + c.plu + '</td>';
      h += '<td>' + c.sales + '</td><td>' + c.inv + '</td><td>' + c.doh + '</td>';
      h += '<td class="' + sc + '">' + c.status + '</td></tr>';
    });
    h += '</tbody></table></div>';
  }

  // Brand display analysis
  if (d.brands && d.brands.length > 0) {
    h += '<h3 style="color:#0b8043;margin:20px 0 8px;">🟢 Brand 可展示分析</h3>';
    h += '<div class="scroll-wrapper"><table><thead><tr>';
    h += '<th class="green">品牌</th><th class="green">佔銷量%</th><th class="green">SKU數</th><th class="green">可展示SKU(≥2)</th><th class="green">總庫存</th><th class="green">DOH(月)</th>';
    h += '</tr></thead><tbody>';
    d.brands.forEach(b => {
      const dcls = b.display === 0 ? 'val-zero' : (b.display <= 2 ? 'val-low' : 'val-ok');
      h += '<tr><td style="text-align:left;font-weight:600">' + b.brand + '</td>';
      h += '<td>' + b.pct.toFixed(1) + '%</td><td>' + b.sku + '</td>';
      h += '<td class="' + dcls + '">' + b.display + '</td>';
      h += '<td>' + b.inv + '</td><td>' + b.doh + '</td></tr>';
    });
    h += '</tbody></table></div>';
  }

  area.innerHTML = h;
}

function filterStoreTable(q) {
  const rows = document.querySelectorAll('#store-frame-table tbody tr');
  q = q.toLowerCase();
  rows.forEach(r => {
    const c0 = r.cells[0] ? r.cells[0].textContent.toLowerCase() : '';
    const c2 = r.cells[2] ? r.cells[2].textContent.toLowerCase() : '';
    r.style.display = (c0.includes(q) || c2.includes(q)) ? '' : 'none';
  });
}
</script>
</body>
</html>"""

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"    -> HTML report saved: {output_path}")
    return output_path
