"""
store_health_auto.py ─ Fully Automated Store Health Check
Downloads 在庫表, runs analysis, generates report — no user input needed.

Default: Last 30 days for frames, Last 60 days for CL, all stores.
Override via command-line args:
  python store_health_auto.py [--days 15] [--stores 101,102] [--top 10] [--date 2026-04-01]
"""

import os, sys, time, glob, shutil, traceback, argparse
from datetime import datetime, date, timedelta

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

CONFIG = {
    "ID":           "OD12558",
    "PASS":         "Howsiao520!",
    "DOWNLOAD_DIR": BASE_DIR,
    "REPORT_URL":   "https://core.owndays.net/ChoHyoDetails/MainChoHyo",
}

# Defaults
DEFAULT_SALES_DAYS = 30
DEFAULT_CL_SALES_DAYS = 60


# ═══════════════════════════════════════════════════════════════════════════════
# SELENIUM HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def wait_click(driver, by, locator, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, locator)))

def wait_present(driver, by, locator, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, locator)))

def safe_check(driver, element_id):
    cb = driver.find_element(By.ID, element_id)
    if not cb.is_selected():
        cb.click()

def set_date_field(driver, field_id, value, retries=3):
    """Set date field value via JS with retry — field may not be ready after postback."""
    for attempt in range(retries):
        try:
            field = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, field_id)))
            driver.execute_script(
                "var el = document.getElementById(arguments[0]);"
                "if(el){ el.value = arguments[1]; el.dispatchEvent(new Event('change')); }",
                field_id, value)
            # Verify it was set
            actual = driver.execute_script(
                "return document.getElementById(arguments[0]).value", field_id)
            if actual == value:
                return
            time.sleep(1)
        except Exception:
            if attempt < retries - 1:
                print(f"      -> Retry date field ({attempt+1}/{retries})...")
                time.sleep(2)
            else:
                raise

def wait_for_download(download_dir, before_files, timeout=600):
    deadline = time.time() + timeout
    while time.time() < deadline:
        after = set(
            glob.glob(os.path.join(download_dir, "*.xlsx")) +
            glob.glob(os.path.join(download_dir, "*.xls")))
        new_files = [f for f in (after - before_files) if not f.endswith(".crdownload")]
        if new_files:
            return max(new_files, key=os.path.getmtime)
        dl_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        for pat in ["*.xlsx", "*.xls"]:
            for f in glob.glob(os.path.join(dl_folder, pat)):
                if time.time() - os.path.getmtime(f) < timeout + 5:
                    if f not in before_files:
                        return f
        time.sleep(2)
    return None

def navigate_to_report_form(driver):
    print("   -> Navigate to report page...")
    try:
        driver.get(CONFIG["REPORT_URL"])
    except Exception:
        # Page load timeout — page may still be usable
        print("      -> Page load slow, continuing...")
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "ddlChoHyoCategory")))
    time.sleep(1)

    print("   -> Select category...")
    cat_elem = wait_click(driver, By.ID, "ddlChoHyoCategory")
    cat_sel = Select(cat_elem)
    try:
        cat_sel.select_by_visible_text("盤點報表")
        print("      -> OK (by text)")
    except Exception:
        for v in ["3", "4", "5", "6"]:
            try:
                cat_sel.select_by_value(v)
                time.sleep(1)
                if driver.find_elements(By.XPATH, "//*[contains(text(),'庫存表')]"):
                    print(f"      -> OK (fallback value={v})")
                    break
            except Exception:
                continue
    time.sleep(2)

    print("   -> Click report button...")
    btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
        By.XPATH,
        "//input[contains(@value,'庫存表')] | "
        "//a[contains(text(),'庫存表')] | "
        "//button[contains(text(),'庫存表')]")))
    btn.click()
    time.sleep(3)

    # Wait for the form to fully load after postback
    print("   -> Waiting for form to load...")
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "MainContent_lbxCountry")))
    time.sleep(1)
    print("      -> OK")

def set_filters(driver, eday, sales_start, sales_end):
    # Country: Taiwan
    print(f"   -> Select country: Taiwan...")
    try:
        country = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "MainContent_lbxCountry")))
        sel = Select(country)
        try: sel.select_by_value("02")
        except: sel.select_by_visible_text("Taiwan")
        time.sleep(3)
        print("      -> OK")
    except Exception as ex:
        print(f"      -> WARN: country selection failed ({ex})")

    # Region: all
    print("   -> Select region...")
    try:
        region = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "MainContent_lbxArea")))
        opts = Select(region).options
        if opts:
            Select(region).select_by_value(opts[0].get_attribute("value"))
            print(f"      -> OK ({opts[0].text})")
        time.sleep(2)
    except Exception as ex:
        print(f"      -> WARN: region failed ({ex})")

    # Display: JAN
    print("   -> Select display type: JAN...")
    try:
        grp = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "MainContent_lbxGroupType12")))
        Select(grp).select_by_value("2")
        print("      -> OK")
    except Exception as ex:
        print(f"      -> WARN: display type failed ({ex})")

    # Business date — wait for field after postback, retry if stale
    print(f"   -> Set business date: {eday}...")
    try:
        time.sleep(2)  # extra wait for postback to settle
        set_date_field(driver, "MainContent_tbxEigyoDate", eday)
        print("      -> OK")
    except Exception as ex:
        print(f"      -> WARN: business date failed, trying JS fallback...")
        try:
            driver.execute_script(
                "var el = document.getElementById('MainContent_tbxEigyoDate');"
                "if(el){ el.value = arguments[0]; }", eday)
            print("      -> OK (JS fallback)")
        except Exception as ex2:
            print(f"      -> WARN: business date failed ({ex2})")

    # Sales date
    print(f"   -> Enable sales date: {sales_start} ~ {sales_end}...")
    try:
        safe_check(driver, "MainContent_cbxUriageChk")
        time.sleep(1)
        set_date_field(driver, "MainContent_tbxUriageFrom", sales_start)
        set_date_field(driver, "MainContent_tbxUriageTo", sales_end)
        print("      -> OK")
    except Exception as ex:
        print(f"      -> WARN: sales date failed ({ex})")

    # No pivot
    print("   -> Check no-pivot...")
    try:
        safe_check(driver, "MainContent_cbxTatehyoji")
        print("      -> OK")
    except Exception as ex:
        print(f"      -> WARN: pivot checkbox failed ({ex})")

def click_download(driver, download_dir, before_files):
    dl_btn = wait_present(driver, By.ID, "MainContent_btnExcel")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", dl_btn)
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", dl_btn)
    return wait_for_download(download_dir, before_files)


# ═══════════════════════════════════════════════════════════════════════════════
# DOWNLOAD + ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════

def download_all(base_date, sales_days, cl_sales_days):
    """Download both frame and CL inventory files."""
    download_dir = CONFIG["DOWNLOAD_DIR"]

    downloads = []
    frame_start = base_date - timedelta(days=sales_days - 1)
    downloads.append({
        'label': f'frames_{sales_days}d',
        'eday': str(base_date),
        'sales_start': str(frame_start),
        'sales_end': str(base_date),
        'filename': f'StoreHealth_frames_{sales_days}d_eday{base_date}.xlsx',
    })

    if cl_sales_days != sales_days:
        cl_start = base_date - timedelta(days=cl_sales_days - 1)
        downloads.append({
            'label': f'CL_{cl_sales_days}d',
            'eday': str(base_date),
            'sales_start': str(cl_start),
            'sales_end': str(base_date),
            'filename': f'StoreHealth_CL_{cl_sales_days}d_eday{base_date}.xlsx',
        })

    print(f"\n  Auto-downloading {len(downloads)} file(s)...")
    for d in downloads:
        print(f"    {d['label']}: {d['sales_start']} ~ {d['sales_end']}")

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options)
    # Large reports can take a long time to generate server-side
    driver.set_page_load_timeout(300)  # 5 min page load timeout

    result = {}

    try:
        print("\n  [Login]...")
        driver.get("https://core.owndays.net/Account/Login")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "tbxUsername"))
        ).send_keys(CONFIG["ID"])
        driver.find_element(By.NAME, "tbxPassword").send_keys(CONFIG["PASS"])
        driver.find_element(By.NAME, "tbxPassword").send_keys(Keys.RETURN)
        WebDriverWait(driver, 15).until(lambda d: "Login" not in d.current_url)
        print("    -> OK")

        MAX_RETRIES = 2
        for idx, dl in enumerate(downloads):
            success = False
            for attempt in range(MAX_RETRIES + 1):
                try:
                    label = f"[{idx+1}/{len(downloads)}] {dl['label']}"
                    if attempt > 0:
                        label += f" (retry {attempt}/{MAX_RETRIES})"
                    print(f"\n  {label}...")

                    navigate_to_report_form(driver)
                    set_filters(driver, dl['eday'], dl['sales_start'], dl['sales_end'])

                    before = set(
                        glob.glob(os.path.join(download_dir, "*.xlsx")) +
                        glob.glob(os.path.join(download_dir, "*.xls")) +
                        glob.glob(os.path.join(os.path.expanduser("~"), "Downloads", "*.xlsx")) +
                        glob.glob(os.path.join(os.path.expanduser("~"), "Downloads", "*.xls")))

                    filepath = click_download(driver, download_dir, before)
                    if filepath:
                        new_path = os.path.join(download_dir, dl['filename'])
                        shutil.move(filepath, new_path)
                        result[dl['label']] = new_path
                        print(f"    -> OK: {dl['filename']} ({os.path.getsize(new_path):,} bytes)")
                        success = True
                        break
                    else:
                        print(f"    -> FAILED: download timeout (waited 10 min)")

                except Exception as ex:
                    print(f"    -> ERROR: {type(ex).__name__}: {str(ex)[:120]}")

                if attempt < MAX_RETRIES:
                    print(f"    -> Retrying in 5 seconds...")
                    time.sleep(5)

            if not success:
                print(f"    -> GAVE UP after {MAX_RETRIES + 1} attempts")

            time.sleep(2)

    except Exception:
        traceback.print_exc()
    finally:
        driver.quit()

    return result


def main():
    parser = argparse.ArgumentParser(description='Store Health Auto Check')
    parser.add_argument('--days', type=int, default=DEFAULT_SALES_DAYS,
                        help=f'Frame sales period in days (default: {DEFAULT_SALES_DAYS})')
    parser.add_argument('--cl-days', type=int, default=DEFAULT_CL_SALES_DAYS,
                        help=f'CL sales period in days (default: {DEFAULT_CL_SALES_DAYS})')
    parser.add_argument('--date', type=str, default=None,
                        help='Base date YYYY-MM-DD (default: yesterday)')
    parser.add_argument('--stores', type=str, default=None,
                        help='Comma-separated store name keywords')
    parser.add_argument('--top', type=int, default=None,
                        help='Only top N stores by sales')
    parser.add_argument('--skip-download', action='store_true',
                        help='Skip download, use existing files')
    args = parser.parse_args()

    base_date = datetime.strptime(args.date, "%Y-%m-%d").date() if args.date else date.today() - timedelta(days=1)
    sales_days = args.days
    cl_sales_days = args.cl_days
    store_filter = [s.strip() for s in args.stores.split(',')] if args.stores else None
    top_n_stores = args.top

    print("=" * 60)
    print("  Store Health Auto Check")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Base date: {base_date}")
    print(f"  Frame period: {sales_days}d | CL period: {cl_sales_days}d")
    if store_filter:
        print(f"  Store filter: {store_filter}")
    if top_n_stores:
        print(f"  Top N stores: {top_n_stores}")
    print("=" * 60)

    frame_tag = f'frames_{sales_days}d'
    cl_tag = f'CL_{cl_sales_days}d'
    frame_file = os.path.join(BASE_DIR, f'StoreHealth_{frame_tag}_eday{base_date}.xlsx')
    cl_file = os.path.join(BASE_DIR, f'StoreHealth_{cl_tag}_eday{base_date}.xlsx')

    if args.skip_download and os.path.exists(frame_file):
        print(f"\n  Using existing files...")
        downloaded = {frame_tag: frame_file}
        if os.path.exists(cl_file):
            downloaded[cl_tag] = cl_file
    else:
        downloaded = download_all(base_date, sales_days, cl_sales_days)

    if frame_tag not in downloaded:
        print("\n  ERROR: Frame file not available!")
        sys.exit(1)

    from store_health_core import run_analysis, generate_report
    from store_health_html import generate_html_report

    results = run_analysis(
        frame_file=downloaded[frame_tag],
        cl_file=downloaded.get(cl_tag, None),
        sales_days=sales_days,
        cl_sales_days=cl_sales_days,
        store_filter=store_filter,
        top_n_stores=top_n_stores,
        base_date=base_date,
    )

    if results is None:
        print("\n  ERROR: Analysis failed!")
        sys.exit(1)

    output_name = f"StoreHealth_{base_date}_{sales_days}d.xlsx"
    output_path = os.path.join(BASE_DIR, output_name)
    generate_report(results, output_path)

    # AI Insight + Anomaly Alerts
    insight_data = None
    try:
        from store_health_insight import run_insight_and_alerts
        insight_data = run_insight_and_alerts(results, base_date=str(base_date))
    except Exception as e:
        print(f"\n[INSIGHT] Error: {e} (continuing without insight)")

    # Inject insight into results for HTML
    if insight_data and insight_data.get('insight'):
        results['ai_insight'] = insight_data['insight']
    if insight_data and insight_data.get('alerts'):
        results['ai_alerts'] = insight_data['alerts']

    # Generate HTML dashboard
    html_name = f"StoreHealth_{base_date}_{sales_days}d.html"
    html_path = os.path.join(BASE_DIR, html_name)
    generate_html_report(results, html_path)

    # Also save as index.html for GitHub Pages (always overwrite latest)
    index_html = os.path.join(BASE_DIR, "docs", "index.html")
    os.makedirs(os.path.join(BASE_DIR, "docs"), exist_ok=True)
    generate_html_report(results, index_html)

    print(f"\n{'=' * 60}")
    print(f"  Done!")
    print(f"  Excel: {output_name}")
    print(f"  HTML:  {html_name}")
    print(f"  GitHub Pages: docs/index.html")
    if insight_data and insight_data.get('alerts'):
        print(f"  Alerts: {len(insight_data['alerts'])} anomalies detected")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
