"""
store_health_interactive.py ─ Semi-Automated Store Health Check
Downloads 在庫表 automatically, then asks user for analysis parameters.

Usage: python store_health_interactive.py
"""

import os, sys, time, glob, shutil, traceback
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


# ═══════════════════════════════════════════════════════════════════════════════
# SELENIUM HELPERS (shared with inventory_download_4periods.py)
# ═══════════════════════════════════════════════════════════════════════════════

def wait_click(driver, by, locator, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, locator)))

def wait_present(driver, by, locator, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, locator)))

def safe_check(driver, element_id):
    cb = driver.find_element(By.ID, element_id)
    if not cb.is_selected():
        cb.click()

def set_date_field(driver, field_id, value):
    field = wait_present(driver, By.ID, field_id)
    driver.execute_script("arguments[0].value = arguments[1]", field, value)

def wait_for_download(download_dir, before_files, timeout=180):
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
    driver.get(CONFIG["REPORT_URL"])
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "ddlChoHyoCategory")))
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

    # Business date — wait for field after postback
    print(f"   -> Set business date: {eday}...")
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "MainContent_tbxEigyoDate")))
        set_date_field(driver, "MainContent_tbxEigyoDate", eday)
        print("      -> OK")
    except Exception as ex:
        print(f"      -> WARN: business date failed ({ex})")

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
# USER INPUT
# ═══════════════════════════════════════════════════════════════════════════════

def ask_params():
    print("\n" + "=" * 60)
    print("  Store Health Check - Interactive Mode")
    print("=" * 60)

    today = date.today()
    yesterday = today - timedelta(days=1)

    # Base date
    raw = input(f"\n  營業日 (base date) [{yesterday}]: ").strip()
    base_date = datetime.strptime(raw, "%Y-%m-%d").date() if raw else yesterday

    # Sales period for frames
    raw = input(f"  鏡框銷售天數 (e.g. 30, 15) [30]: ").strip()
    sales_days = int(raw) if raw else 30

    # CL sales period (fixed at 60)
    cl_sales_days = 60
    print(f"  隱形眼鏡銷售天數: {cl_sales_days} (固定)")

    # Store selection
    print(f"\n  門市選擇:")
    print(f"    1. 全門市")
    print(f"    2. 指定門市 (輸入店名關鍵字，逗號分隔)")
    print(f"    3. Top N 銷量門市")
    choice = input(f"  選擇 [1]: ").strip() or '1'

    store_filter = None
    top_n_stores = None

    if choice == '2':
        raw = input("  店名關鍵字 (逗號分隔): ").strip()
        store_filter = [s.strip() for s in raw.split(',') if s.strip()]
    elif choice == '3':
        raw = input("  Top N 門市 [10]: ").strip()
        top_n_stores = int(raw) if raw else 10

    return {
        'base_date': base_date,
        'sales_days': sales_days,
        'cl_sales_days': cl_sales_days,
        'store_filter': store_filter,
        'top_n_stores': top_n_stores,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# DOWNLOAD
# ═══════════════════════════════════════════════════════════════════════════════

def download_inventory(params):
    """Download 在庫表 files for the specified periods."""
    base_date = params['base_date']
    sales_days = params['sales_days']
    cl_sales_days = params['cl_sales_days']

    downloads_needed = []

    # Frame/sunglasses file
    frame_sales_start = base_date - timedelta(days=sales_days - 1)
    downloads_needed.append({
        'label': f'frames_{sales_days}d',
        'eday': str(base_date),
        'sales_start': str(frame_sales_start),
        'sales_end': str(base_date),
        'filename': f'StoreHealth_frames_{sales_days}d_eday{base_date}.xlsx',
    })

    # CL file (only if different period)
    if cl_sales_days != sales_days:
        cl_sales_start = base_date - timedelta(days=cl_sales_days - 1)
        downloads_needed.append({
            'label': f'CL_{cl_sales_days}d',
            'eday': str(base_date),
            'sales_start': str(cl_sales_start),
            'sales_end': str(base_date),
            'filename': f'StoreHealth_CL_{cl_sales_days}d_eday{base_date}.xlsx',
        })

    print(f"\n  Downloads planned:")
    for d in downloads_needed:
        print(f"    {d['label']}: sales {d['sales_start']} ~ {d['sales_end']}")

    input("\n  Press Enter to start downloading...")

    # Start browser
    download_dir = CONFIG["DOWNLOAD_DIR"]
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

    downloaded = {}

    try:
        # Login
        print("\n  [Login] Logging in...")
        driver.get("https://core.owndays.net/Account/Login")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "tbxUsername"))
        ).send_keys(CONFIG["ID"])
        driver.find_element(By.NAME, "tbxPassword").send_keys(CONFIG["PASS"])
        driver.find_element(By.NAME, "tbxPassword").send_keys(Keys.RETURN)
        WebDriverWait(driver, 15).until(lambda d: "Login" not in d.current_url)
        print("    -> Login OK")

        for idx, dl in enumerate(downloads_needed):
            print(f"\n  [{idx+1}/{len(downloads_needed)}] Downloading {dl['label']}...")
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
                downloaded[dl['label']] = new_path
                size = os.path.getsize(new_path)
                print(f"    -> OK: {dl['filename']} ({size:,} bytes)")
            else:
                print(f"    -> FAILED: timeout!")

            time.sleep(2)

    except Exception:
        traceback.print_exc()
    finally:
        driver.quit()
        print("\n  Browser closed")

    return downloaded


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    params = ask_params()

    # Check if files already exist (skip download)
    frame_tag = f'frames_{params["sales_days"]}d'
    cl_tag = f'CL_{params["cl_sales_days"]}d'
    frame_file = os.path.join(BASE_DIR, f'StoreHealth_{frame_tag}_eday{params["base_date"]}.xlsx')
    cl_file = os.path.join(BASE_DIR, f'StoreHealth_{cl_tag}_eday{params["base_date"]}.xlsx')

    if os.path.exists(frame_file):
        print(f"\n  Found existing file: {os.path.basename(frame_file)}")
        skip = input("  Skip download and use existing? [Y/n]: ").strip().lower()
        if skip != 'n':
            downloaded = {frame_tag: frame_file}
            if os.path.exists(cl_file):
                downloaded[cl_tag] = cl_file
        else:
            downloaded = download_inventory(params)
    else:
        downloaded = download_inventory(params)

    if frame_tag not in downloaded:
        print("\n  ERROR: Frame file not available. Cannot proceed.")
        input("Press Enter to close...")
        return

    # Run analysis
    from store_health_core import run_analysis, generate_report

    frame_path = downloaded[frame_tag]
    cl_path = downloaded.get(cl_tag, None)

    # If CL period same as frame period, CL will be extracted from same file
    if cl_path is None and params['cl_sales_days'] == params['sales_days']:
        cl_path = None  # core will extract from same file

    results = run_analysis(
        frame_file=frame_path,
        cl_file=cl_path,
        sales_days=params['sales_days'],
        cl_sales_days=params['cl_sales_days'],
        store_filter=params.get('store_filter'),
        top_n_stores=params.get('top_n_stores'),
    )

    if results is None:
        print("\n  ERROR: Analysis failed.")
        input("Press Enter to close...")
        return

    # Generate report
    output_name = f"StoreHealth_{params['base_date']}_{params['sales_days']}d.xlsx"
    output_path = os.path.join(BASE_DIR, output_name)
    generate_report(results, output_path)

    print(f"\n{'=' * 60}")
    print(f"  Done! Report: {output_name}")
    print(f"{'=' * 60}")
    input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
