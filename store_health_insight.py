"""
store_health_insight.py — AI-powered inventory health insights + LINE anomaly alerts.

Two main features:
  1. Claude API: Generate Chinese insight summary from analysis results
  2. LINE Bot: Push anomaly alerts when significant issues detected

Setup:
  Set environment variables (or put in .env file):
    ANTHROPIC_API_KEY=sk-ant-...
    LINE_CHANNEL_TOKEN=your_line_channel_access_token
    LINE_TARGET_ID=group_or_user_id  (C... for group, U... for user)

  Or pass via config dict to each function.
"""

import os, json
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIG (loaded from env or .env file)
# ═══════════════════════════════════════════════════════════════════════════════

def _load_env():
    """Load .env file if it exists (simple key=value parser)."""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    if os.path.exists(env_path):
        with open(env_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    k, v = line.split('=', 1)
                    os.environ.setdefault(k.strip(), v.strip())

_load_env()

def get_config():
    return {
        'anthropic_api_key': os.environ.get('ANTHROPIC_API_KEY', ''),
        'line_channel_token': os.environ.get('LINE_CHANNEL_TOKEN', ''),
        'line_target_id': os.environ.get('LINE_TARGET_ID', ''),
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 1. CLAUDE API INSIGHT
# ═══════════════════════════════════════════════════════════════════════════════

def _build_prompt(results):
    """Build a structured prompt from analysis results for Claude."""
    summary = results['summary']
    df_frames = results['df_frames']
    retail_stores = results['retail_stores']
    n_stores = len(retail_stores)

    # Top-N metrics
    tier_lines = []
    for n in [50, 100, 200, 300]:
        tier = df_frames[df_frames['national_rank'] <= n]
        ct = len(tier)
        if ct == 0:
            continue
        oos = tier['is_stockout'].sum()
        low = tier['is_low_stock'].sum()
        avg_doh = tier[tier['doh_days'] < 9999]['doh_months'].mean()
        tier_lines.append(
            f"Top {n}: 缺貨率={oos/ct*100:.1f}%, 緊張率={low/ct*100:.1f}%, "
            f"不足率={(oos+low)/ct*100:.1f}%, Avg DOH={avg_doh:.2f}月"
        )

    # Per-store worst 5
    store_pcts = []
    for store in retail_stores:
        sf = df_frames[df_frames['store_name'] == store].sort_values('national_rank').head(50)
        if len(sf) == 0:
            continue
        pct = (sf['is_stockout'].sum() + sf['is_low_stock'].sum()) / len(sf) * 100
        store_pcts.append((store, pct))
    store_pcts.sort(key=lambda x: x[1], reverse=True)

    worst_stores = '\n'.join(f"  {s}: Top50不足率 {p:.0f}%" for s, p in store_pcts[:5])
    best_stores = '\n'.join(f"  {s}: Top50不足率 {p:.0f}%" for s, p in store_pcts[-3:])

    # Top 5 most critical SKUs (highest national sales but stockout in many stores)
    sku_agg = df_frames.groupby('PLU').agg(
        品番=('品番', 'first'),
        national_rank=('national_rank', 'first'),
        total_sales=('sales', 'sum'),
        total_inv=('inventory', 'sum'),
        oos_stores=('is_stockout', 'sum'),
        n_stores=('store_name', 'nunique'),
    ).sort_values('national_rank')
    critical = sku_agg[sku_agg['oos_stores'] >= sku_agg['n_stores'] * 0.3].head(5)

    critical_lines = []
    for _, r in critical.iterrows():
        critical_lines.append(
            f"  #{int(r['national_rank'])} {r['品番']}: "
            f"銷量={int(r['total_sales'])}, 庫存={int(r['total_inv'])}, "
            f"缺貨門市={int(r['oos_stores'])}/{int(r['n_stores'])}家"
        )

    base_date = summary.get('base_date', datetime.now().strftime('%Y-%m-%d'))

    prompt = f"""你是 OWNDAYS Taiwan 的庫存分析專家。以下是 {base_date} 的全台門市庫存健康檢查結果。
請用繁體中文寫一段 150-200 字的洞察摘要，重點標出最需要關注的 2-3 件事。
語氣專業但簡潔，像寫給區域主管的早報。不要用 bullet point，用段落式。

=== 基本資訊 ===
分析門市數: {n_stores}
鏡框SKU數: {summary['total_frame_skus']}
銷售期間: {summary['sales_days']}天

=== 全國 Top N 庫存不足率 ===
{chr(10).join(tier_lines)}

=== 最差5家門市 (Top50 SKU不足率) ===
{worst_stores}

=== 最佳3家門市 ===
{best_stores}

=== 最需關注的缺貨SKU (>30%門市缺貨) ===
{chr(10).join(critical_lines) if critical_lines else '（無嚴重缺貨SKU）'}

=== 其他 ===
隱形眼鏡缺貨率: {summary['cl_stockout_pct']:.1f}%
無用庫存筆數: {summary['dead_stock_count']}
"""
    return prompt


def generate_insight(results, api_key=None):
    """Call Claude Sonnet API to generate insight summary.
    Returns insight text string, or None if API call fails.
    """
    if api_key is None:
        api_key = get_config()['anthropic_api_key']

    if not api_key:
        print("[INSIGHT] No ANTHROPIC_API_KEY set, skipping AI insight.")
        return None

    prompt = _build_prompt(results)

    try:
        from urllib import request as urllib_request
        import json as _json

        body = _json.dumps({
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 1024,
            "messages": [{"role": "user", "content": prompt}],
        }).encode('utf-8')

        req = urllib_request.Request(
            "https://api.anthropic.com/v1/messages",
            data=body,
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
            method="POST"
        )

        resp = urllib_request.urlopen(req, timeout=60)
        result = _json.loads(resp.read().decode('utf-8'))
        insight = result['content'][0]['text']

        print(f"[INSIGHT] Generated {len(insight)} chars insight.")
        return insight

    except Exception as e:
        print(f"[INSIGHT] API call failed: {e}")
        return None


# ═══════════════════════════════════════════════════════════════════════════════
# 2. ANOMALY DETECTION
# ═══════════════════════════════════════════════════════════════════════════════

def detect_anomalies(results, thresholds=None):
    """Detect significant inventory anomalies.
    Returns list of alert dicts: [{'level': 'critical'|'warning', 'message': str}, ...]
    """
    if thresholds is None:
        thresholds = {
            'top50_shortage_critical': 50,   # Top50不足率 > 50% = critical
            'top50_shortage_warning': 40,     # > 40% = warning
            'store_crisis_pct': 70,           # 門市 Top50不足率 > 70%
            'store_crisis_count': 5,          # 超過5家達到crisis = alert
            'sku_oos_store_ratio': 0.5,       # SKU在>50%門市缺貨 = alert
            'sku_oos_min_rank': 50,           # 只看Top50 SKU
        }

    df_frames = results['df_frames']
    retail_stores = results['retail_stores']
    summary = results['summary']
    n_stores = len(retail_stores)
    alerts = []

    # Check 1: National Top50 shortage rate
    top50 = df_frames[df_frames['national_rank'] <= 50]
    if len(top50) > 0:
        oos = top50['is_stockout'].sum()
        low = top50['is_low_stock'].sum()
        shortage_pct = (oos + low) / len(top50) * 100

        if shortage_pct > thresholds['top50_shortage_critical']:
            alerts.append({
                'level': 'critical',
                'category': 'shortage',
                'message': f'🔴 全國 Top 50 庫存不足率達 {shortage_pct:.0f}%，超過 {thresholds["top50_shortage_critical"]}% 警戒線！',
            })
        elif shortage_pct > thresholds['top50_shortage_warning']:
            alerts.append({
                'level': 'warning',
                'category': 'shortage',
                'message': f'🟡 全國 Top 50 庫存不足率 {shortage_pct:.0f}%，接近警戒水位',
            })

    # Check 2: Store crisis count
    crisis_stores = []
    for store in retail_stores:
        sf = df_frames[df_frames['store_name'] == store].sort_values('national_rank').head(50)
        if len(sf) == 0:
            continue
        pct = (sf['is_stockout'].sum() + sf['is_low_stock'].sum()) / len(sf) * 100
        if pct > thresholds['store_crisis_pct']:
            crisis_stores.append((store, pct))

    if len(crisis_stores) >= thresholds['store_crisis_count']:
        crisis_stores.sort(key=lambda x: x[1], reverse=True)
        store_list = '、'.join(f"{s}({p:.0f}%)" for s, p in crisis_stores[:5])
        alerts.append({
            'level': 'critical',
            'category': 'store_crisis',
            'message': f'🔴 {len(crisis_stores)}家門市 Top50不足率超過 {thresholds["store_crisis_pct"]}%：{store_list}',
        })

    # Check 3: Critical SKU (top-ranked SKU out of stock in many stores)
    if len(top50) > 0:
        sku_oos = top50.groupby('PLU').agg(
            品番=('品番', 'first'),
            rank=('national_rank', 'first'),
            oos=('is_stockout', 'sum'),
            n=('store_name', 'nunique'),
        )
        critical_skus = sku_oos[sku_oos['oos'] >= sku_oos['n'] * thresholds['sku_oos_store_ratio']]
        critical_skus = critical_skus.sort_values('rank')

        if not critical_skus.empty:
            sku_list = '、'.join(
                f"#{int(r['rank'])} {r['品番']}({int(r['oos'])}/{int(r['n'])}家缺貨)"
                for _, r in critical_skus.head(3).iterrows()
            )
            alerts.append({
                'level': 'critical' if len(critical_skus) >= 3 else 'warning',
                'category': 'sku_crisis',
                'message': f'🔴 Top 50 中有 {len(critical_skus)} 個SKU超過半數門市缺貨：{sku_list}',
            })

    # Check 4: DOH dropping below 1 month for top SKUs
    if len(top50) > 0:
        valid_doh = top50[top50['doh_days'] < 9999]['doh_months']
        avg_doh = valid_doh.mean() if len(valid_doh) > 0 else 0
        if avg_doh < 1.0:
            alerts.append({
                'level': 'warning',
                'category': 'doh',
                'message': f'🟡 Top 50 平均 DOH 僅 {avg_doh:.1f} 月，庫存水位偏低',
            })

    return alerts


# ═══════════════════════════════════════════════════════════════════════════════
# 3. LINE MESSAGING API PUSH
# ═══════════════════════════════════════════════════════════════════════════════

def send_line_message(message, channel_token=None, target_id=None):
    """Send a push message via LINE Messaging API.
    target_id: starts with 'C' for group, 'U' for user, 'R' for room.
    """
    if channel_token is None:
        channel_token = get_config()['line_channel_token']
    if target_id is None:
        target_id = get_config()['line_target_id']

    if not channel_token or not target_id:
        print("[LINE] No LINE_CHANNEL_TOKEN or LINE_TARGET_ID set, skipping.")
        return False

    try:
        from urllib import request as urllib_request
        import json as _json

        body = _json.dumps({
            "to": target_id,
            "messages": [{"type": "text", "text": message}],
        }).encode('utf-8')

        req = urllib_request.Request(
            "https://api.line.me/v2/bot/message/push",
            data=body,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {channel_token}",
            },
            method="POST"
        )

        resp = urllib_request.urlopen(req, timeout=30)
        print(f"[LINE] Message sent ({len(message)} chars)")
        return True

    except Exception as e:
        print(f"[LINE] Push failed: {e}")
        return False


def format_alert_message(alerts, base_date=None):
    """Format anomaly alerts into a LINE message string."""
    if not alerts:
        return None

    date_str = base_date or datetime.now().strftime('%Y-%m-%d')
    header = f"⚠️ OWNDAYS 庫存異常警報\n📅 {date_str}\n{'─' * 20}\n"

    critical = [a for a in alerts if a['level'] == 'critical']
    warning = [a for a in alerts if a['level'] == 'warning']

    lines = [header]
    for a in critical:
        lines.append(a['message'])
    for a in warning:
        lines.append(a['message'])

    lines.append(f"\n📊 詳細請查看 Dashboard")

    return '\n'.join(lines)


# ═══════════════════════════════════════════════════════════════════════════════
# 4. MAIN PIPELINE (called from store_health_auto.py)
# ═══════════════════════════════════════════════════════════════════════════════

def run_insight_and_alerts(results, base_date=None):
    """Run full insight + alert pipeline.
    Returns dict with 'insight' text and 'alerts' list.
    """
    config = get_config()
    base_date_str = base_date or results['summary'].get('base_date', '')

    print(f"\n{'=' * 60}")
    print(f"  AI Insight & Anomaly Detection")
    print(f"{'=' * 60}")

    # 1. Generate Claude insight
    insight = None
    if config['anthropic_api_key']:
        print("\n[1] Generating AI insight (Claude Sonnet)...")
        insight = generate_insight(results, config['anthropic_api_key'])
        if insight:
            print(f"    ✅ Insight generated ({len(insight)} chars)")
            print(f"\n    --- AI 摘要 ---")
            print(f"    {insight[:200]}...")
    else:
        print("\n[1] Skipping AI insight (no ANTHROPIC_API_KEY)")

    # 2. Detect anomalies
    print("\n[2] Running anomaly detection...")
    alerts = detect_anomalies(results)
    if alerts:
        print(f"    ⚠️ Found {len(alerts)} anomalies:")
        for a in alerts:
            print(f"      [{a['level'].upper()}] {a['message']}")
    else:
        print(f"    ✅ No significant anomalies detected")

    # 3. Send LINE alert (only if anomalies found)
    if alerts and config['line_channel_token'] and config['line_target_id']:
        print("\n[3] Sending LINE alert...")
        msg = format_alert_message(alerts, base_date_str)
        if msg:
            send_line_message(msg, config['line_channel_token'], config['line_target_id'])
    elif alerts:
        print("\n[3] Anomalies found but LINE not configured (set LINE_CHANNEL_TOKEN & LINE_TARGET_ID)")
    else:
        print("\n[3] No anomalies → no LINE push")

    return {
        'insight': insight,
        'alerts': alerts,
    }
