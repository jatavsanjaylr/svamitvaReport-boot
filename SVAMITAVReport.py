import os
import sys
import time
import glob
import base64
import logging
import sqlite3
import warnings
import traceback
from datetime import datetime, timedelta
from io import StringIO
from urllib.parse import quote
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Pandas Future/Deprecation Warnings ko silent karna
warnings.filterwarnings('ignore', category=UserWarning)
warnings.filterwarnings('ignore', category=FutureWarning)

# PowerPoint PPTX Library
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE

    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False
    print("\n[WARNING] 'python-pptx' install nahi hai. PPT skip hogi (pip install python-pptx chalayein).\n")

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# =========================================================
# 1. Dynamic Drive aur Database Folder Setup
# =========================================================
if os.path.exists("D:\\"):
    TARGET_DRIVE = "D:\\"
else:
    TARGET_DRIVE = "C:\\"

MAIN_FOLDER = os.path.join(TARGET_DRIVE, "SVAMITVA_Ashoknagar")
DB_PATH = os.path.join(MAIN_FOLDER, "svamitva_master.db")

run_timestamp = datetime.now().strftime("%d-%m-%Y %I.%M %p")
PROGRESS_SUBFOLDER = f"SVAMITVA Registry Progress {run_timestamp}"

OUTPUT_DIR = os.path.join(MAIN_FOLDER, PROGRESS_SUBFOLDER)
BASE_DOWNLOAD_DIR = os.path.join(MAIN_FOLDER, "temp_downloads")
LOG_DIR = os.path.join(MAIN_FOLDER, "logs")
WHATSAPP_SESSION_DIR = os.path.join(TARGET_DRIVE, "whatsapp_session")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(WHATSAPP_SESSION_DIR, exist_ok=True)

BASE_DIR = OUTPUT_DIR

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, "automation.log"), encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

print(f"\n=======================================================")
print(f"डेटाबेस पाथ: {DB_PATH}")
print(f"आउटपुट फ़ोल्डर: {OUTPUT_DIR}")
print(f"=======================================================\n")

MAIN_DASHBOARD_URL = "https://services.mpigr.gov.in/svamitva_dashboard/?view=district"
TARGET_DISTRICT = "अशोकनगर"

GWALIOR_DIVISION_DISTRICTS = ["ग्वालियर", "अशोकनगर", "गुना", "शिवपुरी", "दतिया", "gwalior", "ashoknagar", "guna",
                              "shivpuri", "datia"]

TEHSILS_DATA = {
    "अशोकनगर": "https://services.mpigr.gov.in/svamitva_dashboard/village/%E0%A4%85%E0%A4%B6%E0%A5%8B%E0%A4%95%E0%A4%A8%E0%A4%97%E0%A4%B0/%E0%A4%85%E0%A4%B6%E0%A5%8B%E0%A4%95%E0%A4%A8%E0%A4%97%E0%A4%B0",
    "ईसागढ": "https://services.mpigr.gov.in/svamitva_dashboard/village/%E0%A4%85%E0%A4%B6%E0%A5%8B%E0%A4%95%E0%A4%A8%E0%A4%97%E0%A4%B0/%E0%A4%88%E0%A4%B8%E0%A4%BE%E0%A4%97%E0%A4%A2",
    "चन्‍देरी": "https://services.mpigr.gov.in/svamitva_dashboard/village/%E0%A4%85%E0%A4%B6%E0%A5%8B%E0%A4%95%E0%A4%A8%E0%A4%97%E0%A4%B0/%E0%A4%9A%E0%A4%A8%E0%A5%8D%E0%A4%A6%E0%A5%87%E0%A4%B0%E0%A5%80",
    "नईसरांय": "https://services.mpigr.gov.in/svamitva_dashboard/village/%E0%A4%85%E0%A4%B6%E0%A5%8B%E0%A4%95%E0%A4%A8%E0%A4%97%E0%A4%B0/%E0%A4%A8%E0%A4%88%E0%A4%B8%E0%A4%B0%E0%A4%BE%E0%A4%AF",
    "पिपरई": "https://services.mpigr.gov.in/svamitva_dashboard/village/%E0%A4%85%E0%A4%B6%E0%A5%8B%E0%A4%95%E0%A4%A8%E0%A4%97%E0%A4%B0/%E0%A4%AA%E0%A4%BF%E0%A4%AA%E0%A4%B0%E0%A4%88",
    "बहादुरपुर": "https://services.mpigr.gov.in/svamitva_dashboard/village/%E0%A4%85%E0%A4%B6%E0%A5%8B%E0%A4%95%E0%A4%A8%E0%A4%97%E0%A4%B0/%E0%A4%AC%E0%A4%B9%E0%A4%BE%E0%A4%A6%E0%A5%81%E0%A4%B0%E0%A4%AA%E0%A5%81%E0%A4%B0",
    "मुंगावली": "https://services.mpigr.gov.in/svamitva_dashboard/village/%E0%A4%85%E0%A4%B6%E0%A5%8B%E0%A4%95%E0%A4%A8%E0%A4%97%E0%A4%B0/%E0%A4%AE%E0%A5%81%E0%A4%82%E0%A4%97%E0%A4%BE%E0%A4%B5%E0%A4%B2%E0%A5%80",
    "शाढोरा": "https://services.mpigr.gov.in/svamitva_dashboard/village/%E0%A4%85%E0%A4%B6%E0%A5%8B%E0%A4%95%E0%A4%A8%E0%A4%97%E0%A4%B0/%E0%A4%B6%E0%A4%BE%E0%A4%A2%E0%A5%8B%E0%A4%B0%E0%A4%BE"
}


def get_current_time_str():
    return datetime.now().strftime("%d/%m/%Y %I:%M %p")


# =========================================================
# 2. Database Schema (SQLite)
# =========================================================
def init_database():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
                   CREATE TABLE IF NOT EXISTS tehsil_daily_history
                   (
                       id
                       INTEGER
                       PRIMARY
                       KEY
                       AUTOINCREMENT,
                       run_date
                       TEXT,
                       run_datetime
                       TEXT,
                       tehsil_name
                       TEXT,
                       total_patwaris
                       INTEGER,
                       total_villages
                       INTEGER,
                       total_plots
                       INTEGER,
                       initiated
                       INTEGER,
                       verification
                       INTEGER,
                       submitted
                       INTEGER,
                       completed
                       INTEGER,
                       percentage
                       REAL
                   )
                   """)
    cursor.execute("""
                   CREATE TABLE IF NOT EXISTS patwari_daily_history
                   (
                       id
                       INTEGER
                       PRIMARY
                       KEY
                       AUTOINCREMENT,
                       run_date
                       TEXT,
                       run_datetime
                       TEXT,
                       tehsil_name
                       TEXT,
                       halka_name
                       TEXT,
                       total_villages
                       INTEGER,
                       total_plots
                       INTEGER,
                       completed
                       INTEGER,
                       percentage
                       REAL
                   )
                   """)
    conn.commit()
    conn.close()


# =========================================================
# 3. Safe Save Utility
# =========================================================
def safe_save_excel(obj, file_path):
    try:
        if isinstance(obj, Workbook):
            obj.save(file_path)
        else:
            obj.to_excel(file_path, index=False)
        return file_path
    except PermissionError:
        base, ext = os.path.splitext(file_path)
        alt_path = f"{base}_{int(time.time())}{ext}"
        logging.warning(f"File locked! Naye path par save kiya gaya: {os.path.basename(alt_path)}")
        if isinstance(obj, Workbook):
            obj.save(alt_path)
        else:
            obj.to_excel(alt_path, index=False)
        return alt_path


# =========================================================
# 4. Data Extractor
# =========================================================
def get_tehsil_dataframe(driver, tehsil_name):
    before_files = set(glob.glob(os.path.join(BASE_DOWNLOAD_DIR, "*.*")))
    wait = WebDriverWait(driver, 25)

    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//table | //tbody//tr")))
        time.sleep(3)
    except Exception:
        pass

    btn_found = False
    excel_xpaths = [
        "//button[contains(@class, 'buttons-excel')]",
        "//a[contains(@class, 'buttons-excel')]",
        "//button[contains(., 'Excel') or contains(., 'डाउनलोड')]",
        "//a[contains(., 'Excel')]"
    ]
    for xp in excel_xpaths:
        try:
            elem = driver.find_element(By.XPATH, xp)
            driver.execute_script("arguments[0].click();", elem)
            btn_found = True
            break
        except Exception:
            continue

    if btn_found:
        for _ in range(8):
            time.sleep(1)
            current_files = set(glob.glob(os.path.join(BASE_DOWNLOAD_DIR, "*.*")))
            new_files = current_files - before_files
            valid = [f for f in new_files if f.endswith(('.xlsx', '.xls', '.csv')) and not f.endswith('.crdownload')]
            if valid:
                latest = max(valid, key=os.path.getctime)
                try:
                    return pd.read_excel(latest)
                except Exception:
                    return pd.read_csv(latest)

    logging.info(f"[{tehsil_name}] सीधे वेबपेज से टेबल पढ़ी जा रही है...")
    try:
        tables = pd.read_html(StringIO(driver.page_source))
        if tables:
            return max(tables, key=lambda t: len(t.columns) * len(t))
    except Exception as e:
        logging.error(f"[{tehsil_name}] टेबल निष्कर्षण त्रुटि: {e}")

    return None


def find_column_by_keywords(columns, keywords):
    for col in columns:
        c_str = str(col).strip().lower().replace(" ", "").replace("_", "")
        if any(k in c_str for k in keywords):
            return col
    return None


# =========================================================
# 5. Tehsil Data Processing
# =========================================================
def process_tehsil_data(df, tehsil_name):
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = ['_'.join([str(c) for c in col if str(c) and 'Unnamed' not in str(c)]).strip() for col in
                      df.columns]

    drop_patterns = ["ग्राम", "गाँव", "गाव", "village", "lgd", "एल.जी.डी", "एलजीडी", "code", "क्र.", "sr", "s.no"]
    cols_to_drop = []
    patwari_col = None

    for col in df.columns:
        c_str = str(col).strip().lower()
        if any(k in c_str for k in ["पटवारी", "हल्का", "halka", "patwari"]):
            if not patwari_col:
                patwari_col = col
                continue
        if any(k in c_str for k in drop_patterns):
            cols_to_drop.append(col)

    clean_df = df.drop(columns=cols_to_drop, errors='ignore').copy()

    if not patwari_col:
        text_cols = clean_df.select_dtypes(include=['object', 'string']).columns.tolist()
        patwari_col = text_cols[0] if text_cols else clean_df.columns[0]

    for c in clean_df.columns:
        if c != patwari_col:
            clean_df[c] = pd.to_numeric(clean_df[c].astype(str).str.replace('%', '', regex=False).str.strip(),
                                        errors='coerce').fillna(0)

    total_plot_col = find_column_by_keywords(clean_df.columns,
                                             ["कुलप्लॉट", "totalplot", "कुलplots", "totalplots", "कुलकिता", "totalkita",
                                              "totalparcel", "plot"])
    complete_plot_col = find_column_by_keywords(clean_df.columns,
                                                ["पूर्णप्लॉट", "कम्प्लीटप्लॉट", "कम्प्लीट", "पूर्ण", "completedplot",
                                                 "completeplot", "completed"])
    pct_col_name = find_column_by_keywords(clean_df.columns,
                                           ["प्रतिशत", "percent", "%", "pratishat"]) or "प्रगति प्रतिशत"

    num_cols = [c for c in clean_df.select_dtypes(include=['number']).columns if c != pct_col_name]
    agg_dict = {col: 'sum' for col in num_cols}
    village_counts = df.groupby(patwari_col).size().rename("कुल ग्राम")

    grouped_df = clean_df.groupby(patwari_col, as_index=False).agg(agg_dict)
    grouped_df = grouped_df.merge(village_counts, on=patwari_col, how='left')

    if complete_plot_col and total_plot_col:
        grouped_df[pct_col_name] = grouped_df.apply(
            lambda r: round((r[complete_plot_col] / r[total_plot_col] * 100), 2) if r[total_plot_col] > 0 else 0.0,
            axis=1
        )
    else:
        grouped_df[pct_col_name] = 0.0

    if complete_plot_col:
        grouped_df = grouped_df.sort_values(by=[complete_plot_col, pct_col_name], ascending=[False, False]).reset_index(
            drop=True)
    else:
        grouped_df = grouped_df.sort_values(by=patwari_col).reset_index(drop=True)

    grouped_df.insert(0, "तहसील", tehsil_name)
    other_cols = [c for c in grouped_df.columns if c not in ["तहसील", patwari_col, "कुल ग्राम", pct_col_name]]
    ordered_cols = ["तहसील", patwari_col, "कुल ग्राम"] + other_cols + [pct_col_name]
    grouped_df = grouped_df[ordered_cols]

    count_1_to_9 = 0
    count_zero = 0
    if complete_plot_col and complete_plot_col in grouped_df.columns:
        count_1_to_9 = int(((grouped_df[complete_plot_col] > 0) & (grouped_df[complete_plot_col] < 10)).sum())
        count_zero = int((grouped_df[complete_plot_col] == 0).sum())
        zero_patwaris = grouped_df[grouped_df[complete_plot_col] == 0].copy()
    else:
        zero_patwaris = pd.DataFrame()

    tot_patwaris = len(grouped_df)
    tot_villages = grouped_df["कुल ग्राम"].sum()

    total_row = {c: "" for c in grouped_df.columns}
    total_row["तहसील"] = f"महायोग ({tehsil_name})"
    total_row[patwari_col] = f"{tot_patwaris} पटवारी"
    total_row["कुल ग्राम"] = tot_villages

    for col in other_cols:
        total_row[col] = grouped_df[col].sum()

    if complete_plot_col and total_plot_col and total_row[total_plot_col] > 0:
        total_row[pct_col_name] = round((total_row[complete_plot_col] / total_row[total_plot_col] * 100), 2)
    else:
        total_row[pct_col_name] = 0.0

    tehsil_patwari_df = pd.concat([grouped_df, pd.DataFrame([total_row])], ignore_index=True)

    tehsil_goshwara_row = {
        "क्र.": "",
        "तहसील": tehsil_name,
        "कुल पटवारी": tot_patwaris,
        "कुल ग्राम": tot_villages
    }
    for col in other_cols:
        tehsil_goshwara_row[col] = grouped_df[col].sum()
    tehsil_goshwara_row[pct_col_name] = total_row[pct_col_name]

    return grouped_df, tehsil_patwari_df, tehsil_goshwara_row, zero_patwaris, pct_col_name, total_plot_col, complete_plot_col, count_1_to_9, count_zero


# =========================================================
# 6. Formatting Utilities
# =========================================================
def format_excel_file(file_path, pct_col_name, is_zero_report=False):
    try:
        wb = load_workbook(file_path)
    except Exception:
        return
    ws = wb.active

    dt_str = get_current_time_str()
    ws.insert_rows(1)
    max_c = ws.max_column
    ws.cell(row=1, column=max_c, value=f"रिपोर्ट समय: {dt_str}").font = Font(name="Mangal", size=10, italic=True)
    ws.cell(row=1, column=max_c).alignment = Alignment(horizontal="right")

    header_hex = "C00000" if is_zero_report else "1F497D"
    header_fill = PatternFill(start_color=header_hex, end_color=header_hex, fill_type="solid")
    grand_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    header_font = Font(name="Mangal", size=12, bold=True, color="FFFFFF")
    bold_font = Font(name="Mangal", size=11, bold=True)
    regular_font = Font(name="Mangal", size=11)

    thin_border = Border(
        left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6')
    )

    for cell in ws[2]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    max_row = ws.max_row
    header_col_names = [cell.value for cell in ws[2]]

    for row in ws.iter_rows(min_row=3, max_row=max_row):
        is_last = (row[0].row == max_row) and not is_zero_report
        for col_idx, cell in enumerate(row):
            col_name = header_col_names[col_idx]
            cell.border = thin_border
            cell.font = bold_font if is_last else regular_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

            if col_name == pct_col_name and isinstance(cell.value, (int, float)):
                cell.number_format = '0.00"%"'
            elif isinstance(cell.value, float):
                cell.number_format = '0.00'

            if is_last:
                cell.fill = grand_fill

    for col_idx in range(1, len(header_col_names) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(2, max_row + 1))
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 38)

    safe_save_excel(wb, file_path)


def generate_pdf_via_chrome(driver, df, pdf_path, title_text, pct_col_name, is_zero_report=False,
                            footer_notes_html=None, top_performer_box=None):
    dt_str = get_current_time_str()
    headers_html = "".join([f"<th>{col}</th>" for col in df.columns])
    rows_html = ""
    last_idx = len(df) - 1

    for idx, row in df.iterrows():
        is_last = (idx == last_idx) and not is_zero_report
        row_cls = "class='grandtotal-row'" if is_last else ""
        tds = []
        for col_name, val in row.items():
            if pd.isna(val) or val == "":
                display_val = ""
            elif col_name == pct_col_name and isinstance(val, (int, float)):
                display_val = f"{val:.2f}%"
            elif isinstance(val, float):
                display_val = f"{val:.2f}"
            else:
                display_val = str(val)
            tds.append(f"<td>{display_val}</td>")
        rows_html += f"<tr {row_cls}>{''.join(tds)}</tr>"

    theme_color = "#C00000" if is_zero_report else "#1F497D"

    top_performer_html = f"""
    <div style="margin-bottom: 12px; padding: 10px 14px; background-color: #E2EFDA; border: 1.5px solid #385723; border-radius: 5px; font-size: 11pt; color: #1e4620;">
        {top_performer_box}
    </div>
    """ if top_performer_box else ""

    footer_section_html = f"""
    <div style="margin-top: 15px; padding: 12px 16px; background-color: #f7f9fa; border: 1.5px solid #1F497D; border-left: 6px solid #1F497D; border-radius: 5px; font-size: 10.5pt; line-height: 1.6; color: #111;">
        {footer_notes_html}
    </div>
    """ if footer_notes_html else ""

    html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <style>
        @page {{ size: A4 landscape; margin: 8mm 6mm 8mm 6mm; }}
        body {{ font-family: 'Mangal', 'Nirmala UI', sans-serif; margin: 0; padding: 0; color: #111; }}
        table {{ width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 11px; }}
        th, td {{ border: 1px solid #666; padding: 6px 4px; text-align: center; word-wrap: break-word; overflow-wrap: break-word; }}
        th {{ background-color: {theme_color} !important; color: #ffffff !important; font-size: 12px; font-weight: bold; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        tr:nth-child(even) {{ background-color: #f7f9fa; }}
        .grandtotal-row {{ background-color: #d9e1f2 !important; font-size: 12px; font-weight: bold; color: #0c2b53; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    </style>
    </head>
    <body>
        <table style="width: 100%; border: none; margin-bottom: 8px;">
            <tr style="background: none;">
                <td style="border: none; text-align: left; font-size: 15pt; font-weight: bold; color: {theme_color};">{title_text}</td>
                <td style="border: none; text-align: right; font-size: 10pt; color: #444; width: 260px;">दिनांक व समय: {dt_str}</td>
            </tr>
        </table>
        {top_performer_html}
        <table>
            <thead><tr>{headers_html}</tr></thead>
            <tbody>{rows_html}</tbody>
        </table>
        {footer_section_html}
    </body>
    </html>
    """
    temp_html = os.path.join(BASE_DIR, "temp_render.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    driver.get(f"file:///{temp_html}")
    time.sleep(1)

    print_options = {
        'landscape': True, 'displayHeaderFooter': True, 'printBackground': True, 'preferCSSPageSize': True,
        'footerTemplate': '<div style="font-size:10px; text-align:right; width:95%; font-family:Mangal;">पृष्ठ <span class="pageNumber"></span> / <span class="totalPages"></span></div>',
        'headerTemplate': '<div></div>'
    }
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(pdf_path, 'wb') as f:
        f.write(base64.b64decode(result['data']))

    if os.path.exists(temp_html):
        os.remove(temp_html)


# =========================================================
# 7. 10 Se Kam & 0 Progress Summary Report
# =========================================================
def generate_low_progress_report(driver, low_prog_list, output_excel, output_pdf):
    df_low = pd.DataFrame(low_prog_list)
    dt_str = get_current_time_str()

    wb = Workbook()
    ws = wb.active
    ws.title = "प्रगति सारांश"
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    font_title = Font(name="Mangal", size=14, bold=True)
    font_header = Font(name="Mangal", size=12, bold=True)
    font_bold = Font(name="Mangal", size=12, bold=True)
    font_regular = Font(name="Mangal", size=12)
    font_dt = Font(name="Mangal", size=10, italic=True)

    thin_border = Border(
        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
    )

    ws["D1"].value = f"दिनांक: {dt_str}"
    ws["D1"].font = font_dt
    ws["D1"].alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("A2:D2")
    ws["A2"].value = "स्वामित्व योजना में 10 से कम एवं 0 प्रगति वाले पटवारियों की संख्या"
    ws["A2"].font = font_title
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A3:D3")
    ws["A3"].value = "जिनके द्वारा रजिस्ट्री पूर्ण नहीं की गई है"
    ws["A3"].font = font_title
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["स.क्र.", "तहसील का नाम", "10 से कम प्रगति", "0 प्रगति"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=5, column=col_idx, value=h)
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    current_row = 6
    tot_less_10 = 0
    tot_zero = 0

    for idx, r in df_low.iterrows():
        c1 = ws.cell(row=current_row, column=1, value=idx + 1)
        c2 = ws.cell(row=current_row, column=2, value=r["तहसील का नाम"])
        c3 = ws.cell(row=current_row, column=3, value=r["10 से कम प्रगति"])
        c4 = ws.cell(row=current_row, column=4, value=r["0 प्रगति"])

        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c4.alignment = Alignment(horizontal="center", vertical="center")

        for c in [c1, c2, c3, c4]:
            c.font = font_regular
            c.border = thin_border

        tot_less_10 += r["10 से कम प्रगति"]
        tot_zero += r["0 प्रगति"]
        current_row += 1

    c_tot_lbl = ws.cell(row=current_row, column=1, value="कुल")
    c_tot_empty = ws.cell(row=current_row, column=2, value="")
    c_tot_less10 = ws.cell(row=current_row, column=3, value=tot_less_10)
    c_tot_zero = ws.cell(row=current_row, column=4, value=tot_zero)

    c_tot_lbl.alignment = Alignment(horizontal="center", vertical="center")
    c_tot_less10.alignment = Alignment(horizontal="center", vertical="center")
    c_tot_zero.alignment = Alignment(horizontal="center", vertical="center")

    for c in [c_tot_lbl, c_tot_empty, c_tot_less10, c_tot_zero]:
        c.font = font_bold
        c.border = thin_border

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 24
    safe_save_excel(wb, output_excel)

    rows_html = ""
    for idx, r in df_low.iterrows():
        rows_html += f"""
        <tr>
            <td style="text-align: center;">{idx + 1}</td>
            <td style="text-align: left; padding-left: 14px;">{r['तहसील का नाम']}</td>
            <td style="text-align: center;">{r['10 से कम प्रगति']}</td>
            <td style="text-align: center;">{r['0 प्रगति']}</td>
        </tr>
        """
    rows_html += f"""
        <tr style="font-weight: bold; background-color: #f7f9fa;">
            <td style="text-align: center;">कुल</td>
            <td></td>
            <td style="text-align: center;">{tot_less_10}</td>
            <td style="text-align: center;">{tot_zero}</td>
        </tr>
    """

    html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <style>
        @page {{ size: A4 landscape; margin: 12mm 15mm 12mm 15mm; }}
        body {{ font-family: 'Mangal', 'Nirmala UI', sans-serif; color: #000; margin: 0; padding: 0; }}
        .container {{ width: 85%; margin: 0 auto; }}
        .top-dt {{ text-align: right; font-size: 10pt; color: #333; margin-bottom: 6px; }}
        .title-box {{ text-align: center; margin-bottom: 18px; }}
        .title-1 {{ font-size: 17pt; font-weight: bold; margin: 0; }}
        .title-2 {{ font-size: 16pt; font-weight: bold; margin: 6px 0 0 0; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13pt; margin: 0 auto; }}
        th, td {{ border: 1.5px solid #000; padding: 9px 8px; }}
        th {{ background-color: #ffffff; font-weight: bold; text-align: center; font-size: 13.5pt; }}
    </style>
    </head>
    <body>
        <div class="container">
            <div class="top-dt">रिपोर्ट दिनांक व समय: {dt_str}</div>
            <div class="title-box">
                <div class="title-1">स्वामित्व योजना में 10 से कम एवं 0 प्रगति वाले पटवारियों की संख्या</div>
                <div class="title-2">जिनके द्वारा रजिस्ट्री पूर्ण नहीं की गई है</div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 12%;">स.क्र.</th>
                        <th style="width: 44%;">तहसील का नाम</th>
                        <th style="width: 24%;">10 से कम प्रगति</th>
                        <th style="width: 20%;">0 प्रगति</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
    </body>
    </html>
    """
    temp_html = os.path.join(BASE_DIR, "temp_low_prog.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    driver.get(f"file:///{temp_html}")
    time.sleep(1)

    print_options = {'landscape': True, 'printBackground': True, 'preferCSSPageSize': True}
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(output_pdf, 'wb') as f:
        f.write(base64.b64decode(result['data']))

    if os.path.exists(temp_html):
        os.remove(temp_html)


# =========================================================
# 8. State Level Districts Ranking (Only 55 Districts)
# =========================================================
def get_state_districts_data(driver):
    driver.get(MAIN_DASHBOARD_URL)
    wait = WebDriverWait(driver, 25)
    time.sleep(5)

    tab_xpaths = [
        "//a[contains(translate(., 'DISTRICT', 'district'), 'district') and not(contains(translate(., 'TEHSIL', 'tehsil'), 'tehsil'))]",
        "//button[contains(translate(., 'DISTRICT', 'district'), 'district') and not(contains(translate(., 'TEHSIL', 'tehsil'), 'tehsil'))]",
        "//li[contains(@class, 'nav-item')]//a[contains(., 'District') or contains(., 'जिला')]"
    ]
    for xp in tab_xpaths:
        try:
            tab_elem = driver.find_element(By.XPATH, xp)
            driver.execute_script("arguments[0].click();", tab_elem)
            time.sleep(3)
            break
        except Exception:
            continue

    try:
        dropdowns = driver.find_elements(By.XPATH,
                                         "//select[contains(@name, 'length') or contains(@class, 'custom-select')]")
        for dd in dropdowns:
            try:
                driver.execute_script("arguments[0].value = '-1'; arguments[0].dispatchEvent(new Event('change'));", dd)
            except Exception:
                try:
                    driver.execute_script(
                        "arguments[0].value = '100'; arguments[0].dispatchEvent(new Event('change'));", dd)
                except Exception:
                    pass
        time.sleep(3)
    except Exception:
        pass

    tables = pd.read_html(StringIO(driver.page_source))
    if not tables:
        raise ValueError("मुख्य डैशबोर्ड से कोई टेबल प्राप्त नहीं हुई।")

    target_df = None
    for t in tables:
        temp_t = t.copy()
        if isinstance(temp_t.columns, pd.MultiIndex):
            temp_t.columns = ['_'.join([str(c) for c in col if str(c) and 'Unnamed' not in str(c)]).strip() for col in
                              temp_t.columns]

        cols_text = " ".join([str(c).lower() for c in temp_t.columns])
        if ("जिला" in cols_text or "district" in cols_text) and len(temp_t) < 150:
            target_df = temp_t
            break

    if target_df is None:
        valid_tables = [t for t in tables if 45 <= len(t) <= 75]
        if valid_tables:
            target_df = valid_tables[0]
        else:
            non_tehsil = [t for t in tables if len(t) < 200]
            target_df = max(non_tehsil, key=lambda t: len(t)) if non_tehsil else tables[0]

    df = target_df.copy()
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = ['_'.join([str(c) for c in col if str(c) and 'Unnamed' not in str(c)]).strip() for col in
                      df.columns]

    drop_patterns = ["क्र.", "sr", "s.no", "lgd", "code", "action", "तहसील", "tehsil", "taluk", "subdivision"]
    cols_to_drop = [c for c in df.columns if any(k in str(c).lower() for k in drop_patterns)]
    clean_df = df.drop(columns=cols_to_drop, errors='ignore').copy()

    district_col = None
    for c in clean_df.columns:
        c_str = str(c).lower()
        if any(k in c_str for k in ["जिला", "district"]):
            district_col = c
            break
    if not district_col:
        text_cols = clean_df.select_dtypes(include=['object', 'string']).columns.tolist()
        district_col = text_cols[0] if text_cols else clean_df.columns[0]

    clean_df = clean_df[
        ~clean_df[district_col].astype(str).str.contains("कुल|total|महायोग", case=False, na=False)].copy()

    other_text_cols = [c for c in clean_df.select_dtypes(include=['object', 'string']).columns if c != district_col]
    if other_text_cols:
        clean_df = clean_df.drop(columns=other_text_cols, errors='ignore').copy()

    for c in clean_df.columns:
        if c != district_col:
            clean_df[c] = pd.to_numeric(clean_df[c].astype(str).str.replace('%', '', regex=False).str.strip(),
                                        errors='coerce').fillna(0)

    comp_col = None
    tot_plot_col = None
    pct_col = None

    for c in clean_df.columns:
        c_clean = str(c).strip().lower().replace(" ", "").replace("_", "")
        if not comp_col and any(k in c_clean for k in ["कम्प्लीट", "पूर्ण", "completed", "complete"]):
            comp_col = c
        elif not tot_plot_col and any(
                k in c_clean for k in ["कुलप्लॉट", "totalplot", "कुलplots", "totalplots", "कुलकिता", "totalkita"]):
            tot_plot_col = c
        elif not pct_col and any(k in c_clean for k in ["प्रतिशत", "percent", "%", "pratishat"]):
            pct_col = c

    if not pct_col:
        pct_col = "प्रगति प्रतिशत"
        if comp_col and tot_plot_col:
            clean_df[pct_col] = clean_df.apply(
                lambda r: round((r[comp_col] / r[tot_plot_col] * 100), 2) if r[tot_plot_col] > 0 else 0.0,
                axis=1
            )
        else:
            clean_df[pct_col] = 0.0
    else:
        clean_df[pct_col] = clean_df[pct_col].round(2)

    logging.info(f"राज्य स्तरीय तालिका से कुल {len(clean_df)} जिलों का डेटा निकाला गया।")

    return clean_df, district_col, comp_col, pct_col


# --- State Reports Generator with Explicit Total Row in PDF & Excel ---
def generate_district_ranked_report(driver, df_data, district_col, sort_col, sort_label, output_excel, output_pdf,
                                    custom_division_title=None):
    dt_str = get_current_time_str()
    df_sorted = df_data.sort_values(by=sort_col, ascending=False).reset_index(drop=True)
    df_sorted.insert(0, "स.क्र.", range(1, len(df_sorted) + 1))

    total_districts = len(df_sorted)
    ashok_match = df_sorted[df_sorted[district_col].astype(str).str.contains(TARGET_DISTRICT, na=False)]
    ashok_rank = ashok_match["स.क्र."].values[0] if not ashok_match.empty else "N/A"

    if custom_division_title:
        main_header_title = f"{custom_division_title} (प्रतिशत रैंकिंग) | जिला {TARGET_DISTRICT} की रैंक: {ashok_rank}/{total_districts}"
    else:
        main_header_title = f"स्वामित्व योजना राज्य स्तरीय रैंकिंग ({sort_label} अनुसार) | जिला {TARGET_DISTRICT} की रैंक: {ashok_rank}/{total_districts}"

    # Total Row calculation
    total_row = {"स.क्र.": "", district_col: "महायोग (Total)"}
    for col in df_sorted.columns[2:]:
        if "%" in col or "प्रतिशत" in col:
            total_row[col] = round(df_sorted[col].mean(), 2)
        else:
            total_row[col] = int(df_sorted[col].sum())

    df_with_total = pd.concat([df_sorted, pd.DataFrame([total_row])], ignore_index=True)

    # 1. Excel Generation
    wb = Workbook()
    ws = wb.active
    ws.title = "Ranking"
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    title_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    sorted_col_header_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
    yellow_row_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    yellow_col_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    font_title = Font(name="Mangal", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Mangal", size=12, bold=True, color="FFFFFF")
    font_header_yellow = Font(name="Mangal", size=12, bold=True, color="000000")
    font_bold = Font(name="Mangal", size=11, bold=True)
    font_regular = Font(name="Mangal", size=11)
    font_dt = Font(name="Mangal", size=10, italic=True)

    thin_border = Border(
        left=Side(style='thin', color='B0B0B0'), right=Side(style='thin', color='B0B0B0'),
        top=Side(style='thin', color='B0B0B0'), bottom=Side(style='thin', color='B0B0B0')
    )

    num_cols_count = len(df_with_total.columns)
    last_col_letter = get_column_letter(num_cols_count)

    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"].value = f"रिपोर्ट दिनांक व समय: {dt_str}"
    ws["A1"].font = font_dt
    ws["A1"].alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells(f"A2:{last_col_letter}2")
    t_cell = ws["A2"]
    t_cell.value = main_header_title
    t_cell.font = font_title
    t_cell.fill = title_fill
    t_cell.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, col_name in enumerate(df_with_total.columns, 1):
        cell = ws.cell(row=3, column=col_idx, value=str(col_name))
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if col_name == sort_col:
            cell.fill = sorted_col_header_fill
            cell.font = font_header_yellow
        else:
            cell.fill = header_fill
            cell.font = font_header

    current_row = 4
    sort_col_idx = list(df_with_total.columns).index(sort_col) + 1
    total_row_idx = len(df_with_total) - 1

    for idx, row in df_with_total.iterrows():
        is_total_row = (idx == total_row_idx)
        is_ashoknagar = (TARGET_DISTRICT in str(row[district_col])) and not is_total_row

        for col_idx, (col_name, val) in enumerate(row.items(), 1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.border = thin_border

            if col_name == "स.क्र.":
                cell.value = int(val) if val != "" else ""
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == district_col:
                cell.value = str(val)
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif isinstance(val, float):
                cell.value = round(val, 2)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = '0.00"%"' if "%" in col_name or "प्रतिशत" in col_name else '0.00'
            else:
                cell.value = val
                cell.alignment = Alignment(horizontal="center", vertical="center")

            if is_total_row:
                cell.fill = total_fill
                cell.font = font_bold
            elif is_ashoknagar:
                cell.fill = yellow_row_fill
                cell.font = font_bold
            elif col_idx == sort_col_idx:
                cell.fill = yellow_col_fill
                cell.font = font_regular
            else:
                cell.font = font_regular

        current_row += 1

    for col_idx in range(1, num_cols_count + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(3, current_row))
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 36)

    safe_save_excel(wb, output_excel)

    # 2. Chrome HTML to PDF Generation (With explicitly styled Grand Total Row)
    headers_html = "".join(
        [f"<th style='background-color: #FFD966; color: #000;'>{c}</th>" if c == sort_col else f"<th>{c}</th>" for c in
         df_with_total.columns])
    rows_html = ""
    for idx, row in df_with_total.iterrows():
        is_total_row = (idx == total_row_idx)
        is_ashoknagar = (TARGET_DISTRICT in str(row[district_col])) and not is_total_row

        if is_total_row:
            tr_style = "background-color: #D9E1F2; font-weight: bold; font-size: 11px; border-top: 2px solid #000;"
        elif is_ashoknagar:
            tr_style = "background-color: #FFFF00; font-weight: bold;"
        else:
            tr_style = ""

        tds = ""
        for c, val in row.items():
            td_style = ""
            if not is_ashoknagar and not is_total_row and c == sort_col:
                td_style = "background-color: #FFF2CC;"

            if ("%" in str(c) or "प्रतिशत" in str(c)) and isinstance(val, (int, float)):
                display_val = f"{float(val):.2f}%"
            elif isinstance(val, float):
                display_val = f"{val:.2f}"
            else:
                display_val = str(val)

            tds += f"<td style='{td_style}'>{display_val}</td>"
        rows_html += f"<tr style='{tr_style}'>{tds}</tr>"

    html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <style>
        @page {{ size: A4 landscape; margin: 8mm 6mm 8mm 6mm; }}
        body {{ font-family: 'Mangal', 'Nirmala UI', sans-serif; color: #111; margin: 0; padding: 0; }}
        .top-meta {{ text-align: right; font-size: 10pt; color: #444; margin-bottom: 5px; }}
        .header-box {{ background-color: #1F497D; color: #ffffff; text-align: center; padding: 9px 4px; font-size: 14pt; font-weight: bold; margin-bottom: 9px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        table {{ width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 10.5px; }}
        th, td {{ border: 1px solid #777; padding: 5px 3px; text-align: center; word-wrap: break-word; overflow-wrap: break-word; }}
        th {{ background-color: #2F5597; color: #ffffff; font-size: 11.5px; font-weight: bold; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        tr:nth-child(even) {{ background-color: #fafafa; }}
    </style>
    </head>
    <body>
        <div class="top-meta">दिनांक व समय: {dt_str}</div>
        <div class="header-box">{main_header_title}</div>
        <table>
            <thead><tr>{headers_html}</tr></thead>
            <tbody>{rows_html}</tbody>
        </table>
    </body>
    </html>
    """
    temp_html = os.path.join(BASE_DIR, "temp_dist_rank.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    driver.get(f"file:///{temp_html}")
    time.sleep(1)

    print_options = {'landscape': True, 'printBackground': True, 'preferCSSPageSize': True}
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(output_pdf, 'wb') as f:
        f.write(base64.b64decode(result['data']))

    if os.path.exists(temp_html):
        os.remove(temp_html)


# =========================================================
# 9. Database Sync Module
# =========================================================
def save_today_data_to_db(goshwara_df, all_tehsil_patwaris_list, global_tot_col, global_comp_col, pct_col_name):
    today_date = datetime.now().strftime("%Y-%m-%d")
    today_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    init_col = find_column_by_keywords(goshwara_df.columns, ["initiated", "प्रारंभ"])
    veri_col = find_column_by_keywords(goshwara_df.columns, ["verification", "सत्यापन"])
    subm_col = find_column_by_keywords(goshwara_df.columns, ["submitted", "सबमिट"])

    for _, r in goshwara_df.iterrows():
        t_name = str(r["तहसील"]).strip()
        if "महायोग" in t_name or "Grand Total" in t_name:
            continue

        cursor.execute("""
                       INSERT INTO tehsil_daily_history (run_date, run_datetime, tehsil_name, total_patwaris,
                                                         total_villages,
                                                         total_plots, initiated, verification, submitted, completed,
                                                         percentage)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                       """, (
                           today_date, today_dt, t_name, int(r["कुल पटवारी"]), int(r["कुल ग्राम"]),
                           int(r.get(global_tot_col, 0)),
                           int(r.get(init_col, 0)) if init_col else 0,
                           int(r.get(veri_col, 0)) if veri_col else 0,
                           int(r.get(subm_col, 0)) if subm_col else 0,
                           int(r.get(global_comp_col, 0)), float(r.get(pct_col_name, 0.0))
                       ))

    for df_t in all_tehsil_patwaris_list:
        if df_t.empty:
            continue
        t_name = str(df_t["तहसील"].iloc[0]).strip()
        patwari_col = df_t.columns[1]

        for _, r in df_t.iterrows():
            cursor.execute("""
                           INSERT INTO patwari_daily_history (run_date, run_datetime, tehsil_name, halka_name,
                                                              total_villages,
                                                              total_plots, completed, percentage)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                           """, (
                               today_date, today_dt, t_name, str(r[patwari_col]).strip(),
                               int(r["कुल ग्राम"]), int(r.get(global_tot_col, 0)),
                               int(r.get(global_comp_col, 0)), float(r.get(pct_col_name, 0.0))
                           ))

    conn.commit()
    conn.close()
    logging.info("डेटाबेस में डेटा सुरक्षित हो गया।")


# =========================================================
# 10. Database Snapshot Seekers (Morning Baseline & 2-Hour Ago)
# =========================================================
def get_morning_baseline_datetime():
    conn = sqlite3.connect(DB_PATH)
    today_date = datetime.now().strftime("%Y-%m-%d")
    try:
        # Aaj ke din ki sabse pehli run (Earliest run of today)
        df_today = pd.read_sql(
            f"SELECT DISTINCT run_datetime FROM tehsil_daily_history WHERE run_date = '{today_date}' ORDER BY run_datetime ASC LIMIT 1",
            conn)
        if not df_today.empty:
            return df_today['run_datetime'].iloc[0]
        # Agar aaj pehli baar chal raha hai toh pichle din ki aakhiri run
        df_prev = pd.read_sql(
            f"SELECT DISTINCT run_datetime FROM tehsil_daily_history WHERE run_date < '{today_date}' ORDER BY run_datetime DESC LIMIT 1",
            conn)
        if not df_prev.empty:
            return df_prev['run_datetime'].iloc[0]
    except Exception:
        pass
    finally:
        conn.close()
    return None


def get_2hour_ago_datetime():
    conn = sqlite3.connect(DB_PATH)
    try:
        # Lagbhag 1.5 se 3 ghante purani run ko seek karna
        time_threshold = (datetime.now() - timedelta(minutes=100)).strftime("%Y-%m-%d %H:%M:%S")
        df_2h = pd.read_sql(
            f"SELECT DISTINCT run_datetime FROM tehsil_daily_history WHERE run_datetime <= '{time_threshold}' ORDER BY run_datetime DESC LIMIT 1",
            conn)
        if not df_2h.empty:
            return df_2h['run_datetime'].iloc[0]
        # Fallback: Agar exact 2 ghante nahi hai toh 2nd latest run
        df_recent = pd.read_sql(
            "SELECT DISTINCT run_datetime FROM tehsil_daily_history ORDER BY run_datetime DESC LIMIT 2", conn)
        if len(df_recent) >= 2:
            return df_recent['run_datetime'].iloc[1]
    except Exception:
        pass
    finally:
        conn.close()
    return None


# =========================================================
# 10.1 Morning Baseline Comparative Report
# =========================================================
def generate_comparative_growth_report(driver, goshwara_df, global_tot_col, global_comp_col, global_pct_col,
                                       output_excel, output_pdf):
    today_date = datetime.now().strftime("%Y-%m-%d")
    base_dt = get_morning_baseline_datetime()

    prev_map = {}
    if base_dt:
        conn = sqlite3.connect(DB_PATH)
        prev_df = pd.read_sql(
            f"SELECT tehsil_name, completed, percentage FROM tehsil_daily_history WHERE run_datetime = '{base_dt}'",
            conn)
        conn.close()
        for _, r in prev_df.iterrows():
            prev_map[r['tehsil_name']] = {'completed': int(r['completed']), 'percentage': float(r['percentage'])}

    base_label = f"Morning Baseline ({base_dt})" if base_dt else "Baseline (0)"

    comp_rows = []
    tot_prev_comp = 0
    tot_curr_comp = 0
    tot_growth = 0
    tot_plots_all = 0

    for _, r in goshwara_df.iterrows():
        t_name = str(r["तहसील"]).strip()
        if "महायोग" in t_name or "Grand Total" in t_name:
            continue

        c_plots = int(r.get(global_tot_col, 0))
        curr_c = int(r.get(global_comp_col, 0))
        curr_p = float(r.get(global_pct_col, 0.0))

        prev_info = prev_map.get(t_name, {'completed': 0, 'percentage': 0.0})
        prev_c = prev_info['completed']
        prev_p = prev_info['percentage']

        net_growth = curr_c - prev_c
        growth_pct = round(curr_p - prev_p, 2)

        tot_plots_all += c_plots
        tot_prev_comp += prev_c
        tot_curr_comp += curr_c
        tot_growth += net_growth

        comp_rows.append({
            "तहसील": t_name,
            "कुल Plots": c_plots,
            "सुबह पूर्ण Plots": prev_c,
            "वर्तमान पूर्ण Plots": curr_c,
            "दैनिक वृद्धि (Plots)": net_growth,
            "सुबह %": f"{prev_p:.2f}%",
            "वर्तमान %": f"{curr_p:.2f}%",
            "प्रतिशत वृद्धि": f"{growth_pct:+.2f}%"
        })

    comp_df = pd.DataFrame(comp_rows)
    comp_df.insert(0, "स.क्र.", range(1, len(comp_df) + 1))

    wb = Workbook()
    ws = wb.active
    ws.title = "Comparative_Growth"
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    font_title = Font(name="Mangal", size=13, bold=True, color="FFFFFF")
    font_header = Font(name="Mangal", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Mangal", size=10, bold=True)
    font_reg = Font(name="Mangal", size=10)
    font_growth = Font(name="Calibri", size=11, bold=True, color="047857")

    fill_title = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_header = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    fill_growth = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    fill_tot = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    ws.merge_cells("A1:I1")
    ws["A1"].value = f"रिपोर्ट समय: {get_current_time_str()} | तुलना आधार: {base_label}"
    ws["A1"].font = Font(name="Mangal", size=9, italic=True)
    ws["A1"].alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("A2:I2")
    ws["A2"].value = f"स्वामित्व योजना - दैनिक मॉर्निंग बेसलाइन तुलनात्मक प्रगति रिपोर्ट | जिला {TARGET_DISTRICT}"
    ws["A2"].font = font_title
    ws["A2"].fill = fill_title
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 34

    headers = list(comp_df.columns)
    ws.row_dimensions[3].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cur_r = 4
    for _, r in comp_df.iterrows():
        ws.row_dimensions[cur_r].height = 22
        for c_idx, col_name in enumerate(headers, 1):
            cell = ws.cell(row=cur_r, column=c_idx, value=r[col_name])
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if col_name == "तहसील":
                cell.alignment = Alignment(horizontal="left", indent=1, vertical="center")
                cell.font = font_reg
            elif col_name == "दैनिक वृद्धि (Plots)":
                cell.fill = fill_growth
                cell.font = font_growth
            else:
                cell.font = font_reg
        cur_r += 1

    ws.row_dimensions[cur_r].height = 26
    tot_prev_pct = round((tot_prev_comp / tot_plots_all * 100), 2) if tot_plots_all > 0 else 0.0
    tot_curr_pct = round((tot_curr_comp / tot_plots_all * 100), 2) if tot_plots_all > 0 else 0.0

    t_vals = ["", "महायोग (Grand Total)", tot_plots_all, tot_prev_comp, tot_curr_comp, f"+{tot_growth}",
              f"{tot_prev_pct:.2f}%", f"{tot_curr_pct:.2f}%", f"{(tot_curr_pct - tot_prev_pct):+.2f}%"]
    for c_idx, val in enumerate(t_vals, 1):
        cell = ws.cell(row=cur_r, column=c_idx, value=val)
        cell.font = font_bold
        cell.fill = fill_tot
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for c in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18
    ws.column_dimensions["C"].width = 24

    safe_save_excel(wb, output_excel)

    rows_html = ""
    for _, r in comp_df.iterrows():
        growth_badge = f"<span style='background-color:#d1fae5; color:#047857; padding:3px 8px; border-radius:4px; font-weight:bold;'>+{r['दैनिक वृद्धि (Plots)']}</span>" if \
        r['दैनिक वृद्धि (Plots)'] > 0 else f"<span style='color:#666;'>{r['दैनिक वृद्धि (Plots)']}</span>"
        rows_html += f"""
        <tr>
            <td>{r['स.क्र.']}</td>
            <td style="text-align: left; padding-left: 10px;">{r['तहसील']}</td>
            <td>{r['कुल Plots']}</td>
            <td>{r['सुबह पूर्ण Plots']}</td>
            <td>{r['वर्तमान पूर्ण Plots']}</td>
            <td>{growth_badge}</td>
            <td>{r['सुबह %']}</td>
            <td>{r['वर्तमान %']}</td>
            <td style="font-weight:bold; color: #047857;">{r['प्रतिशत वृद्धि']}</td>
        </tr>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <style>
        @page {{ size: A4 landscape; margin: 8mm 10mm 8mm 10mm; }}
        body {{ font-family: 'Mangal', 'Segoe UI', sans-serif; color: #0F172A; margin: 0; padding: 0; }}
        .header-top {{ display: flex; justify-content: space-between; align-items: center; background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%); color: white; padding: 8px 14px; border-radius: 6px; margin-bottom: 6px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        .header-title {{ font-size: 13pt; font-weight: bold; }}
        .header-dt {{ font-size: 9pt; text-align: right; }}
        .compare-badge {{ background-color: #EFF6FF; color: #1E40AF; padding: 6px 12px; font-size: 9.5pt; border-left: 5px solid #3B82F6; border-radius: 4px; margin-bottom: 8px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 9.5pt; text-align: center; }}
        th, td {{ border: 1px solid #CBD5E1; padding: 5px 4px; }}
        th {{ background-color: #1E293B !important; color: white !important; font-size: 10pt; font-weight: bold; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        tr:nth-child(even) {{ background-color: #F8FAFC; }}
        .total-row {{ background-color: #E2E8F0 !important; font-weight: bold; font-size: 10pt; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    </style>
    </head>
    <body>
        <div class="header-top">
            <div class="header-title">स्वामित्व योजना - तहसीलवार दैनिक मॉर्निंग तुलनात्मक प्रगति रिपोर्ट | जिला {TARGET_DISTRICT}</div>
            <div class="header-dt">दिनांक: {today_date}</div>
        </div>
        <div class="compare-badge">
            📅 तुलना आधार: <strong>{base_label}</strong> से आज तक हुई शुद्ध दैनिक प्रगति (Morning Baseline Growth)
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 6%;">स.क्र.</th>
                    <th style="width: 18%;">तहसील</th>
                    <th style="width: 11%;">कुल Plots</th>
                    <th style="width: 11%;">सुबह पूर्ण</th>
                    <th style="width: 11%;">वर्तमान पूर्ण</th>
                    <th style="width: 12%;">दैनिक वृद्धि (Plots)</th>
                    <th style="width: 10%;">सुबह %</th>
                    <th style="width: 10%;">आज %</th>
                    <th style="width: 11%;">नेट वृद्धि %</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
                <tr class="total-row">
                    <td></td>
                    <td>महायोग</td>
                    <td>{tot_plots_all}</td>
                    <td>{tot_prev_comp}</td>
                    <td>{tot_curr_comp}</td>
                    <td>+{tot_growth}</td>
                    <td>{tot_prev_pct:.2f}%</td>
                    <td>{tot_curr_pct:.2f}%</td>
                    <td style="color:#047857;">+{(tot_curr_pct - tot_prev_pct):.2f}%</td>
                </tr>
            </tbody>
        </table>
    </body>
    </html>
    """

    temp_html = os.path.join(BASE_DIR, "temp_compare_morning.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    driver.get(f"file:///{temp_html}")
    time.sleep(1)

    print_options = {
        'landscape': True, 'printBackground': True, 'preferCSSPageSize': True,
        'footerTemplate': '<div style="font-size:9px; text-align:right; width:95%; font-family:Mangal;">पृष्ठ <span class="pageNumber"></span> / <span class="totalPages"></span></div>',
        'headerTemplate': '<div></div>'
    }
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(output_pdf, 'wb') as f:
        f.write(base64.b64decode(result['data']))

    if os.path.exists(temp_html):
        os.remove(temp_html)


# =========================================================
# 10.2 NEW: 2-Hour Ago Comparison Report
# =========================================================
def generate_2hour_ago_comparison_report(driver, goshwara_df, global_tot_col, global_comp_col, global_pct_col,
                                         output_excel, output_pdf):
    dt_str = get_current_time_str()
    two_h_dt = get_2hour_ago_datetime()

    prev_map = {}
    if two_h_dt:
        conn = sqlite3.connect(DB_PATH)
        prev_df = pd.read_sql(
            f"SELECT tehsil_name, completed, percentage FROM tehsil_daily_history WHERE run_datetime = '{two_h_dt}'",
            conn)
        conn.close()
        for _, r in prev_df.iterrows():
            prev_map[r['tehsil_name']] = {'completed': int(r['completed']), 'percentage': float(r['percentage'])}

    base_label = f"Previous 2-Hour Snapshot ({two_h_dt})" if two_h_dt else "Baseline (0)"

    comp_records = []
    tot_prev = 0
    tot_today = 0
    tot_growth = 0
    tot_plots = 0

    for _, r in goshwara_df.iterrows():
        t_name = str(r["तहसील"]).strip()
        if "महायोग" in t_name or "Grand Total" in t_name:
            continue

        c_plots = int(r.get(global_tot_col, 0))
        today_c = int(r.get(global_comp_col, 0))
        today_p = float(r.get(global_pct_col, 0.0))

        prev_info = prev_map.get(t_name, {'completed': 0, 'percentage': 0.0})
        prev_c = prev_info['completed']
        prev_p = prev_info['percentage']

        diff = today_c - prev_c
        p_diff = round(today_p - prev_p, 2)

        tot_plots += c_plots
        tot_prev += prev_c
        tot_today += today_c
        tot_growth += diff

        comp_records.append({
            "तहसील का नाम": t_name,
            "कुल Plots": c_plots,
            "2 घंटे पूर्व पूर्ण": prev_c,
            "वर्तमान पूर्ण": today_c,
            "2 घंटे में वृद्धि (Plots)": diff,
            "2 घंटे पूर्व %": f"{prev_p:.2f}%",
            "वर्तमान %": f"{today_p:.2f}%",
            "वृद्धि %": f"{p_diff:+.2f}%"
        })

    df_result = pd.DataFrame(comp_records)
    df_result = df_result.sort_values(by="2 घंटे में वृद्धि (Plots)", ascending=False).reset_index(drop=True)
    df_result.insert(0, "स.क्र.", range(1, len(df_result) + 1))

    wb = Workbook()
    ws = wb.active
    ws.title = "2Hour_Comparison"
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    font_title = Font(name="Mangal", size=13, bold=True, color="FFFFFF")
    font_header = Font(name="Mangal", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Mangal", size=10, bold=True)
    font_reg = Font(name="Mangal", size=10)
    font_green = Font(name="Calibri", size=11, bold=True, color="047857")

    fill_title = PatternFill(start_color="800080", end_color="800080", fill_type="solid")  # Deep Purple
    fill_header = PatternFill(start_color="9C27B0", end_color="9C27B0", fill_type="solid")
    fill_positive = PatternFill(start_color="F3E5F5", end_color="F3E5F5", fill_type="solid")
    fill_tot = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    ws.merge_cells("A1:I1")
    ws["A1"].value = f"रिपोर्ट समय: {dt_str} | आधार: {base_label}"
    ws["A1"].font = Font(name="Mangal", size=9, italic=True)
    ws["A1"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[1].height = 18

    ws.merge_cells("A2:I2")
    ws["A2"].value = f"स्वामित्व योजना - तहसीलवार 2 घंटे पूर्व से तुलनात्मक समीक्षा रिपोर्ट | जिला {TARGET_DISTRICT}"
    ws["A2"].font = font_title
    ws["A2"].fill = fill_title
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 34

    headers = list(df_result.columns)
    ws.row_dimensions[3].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cur_r = 4
    for _, r in df_result.iterrows():
        ws.row_dimensions[cur_r].height = 22
        for c_idx, col_name in enumerate(headers, 1):
            cell = ws.cell(row=cur_r, column=c_idx, value=r[col_name])
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if col_name == "तहसील का नाम":
                cell.alignment = Alignment(horizontal="left", indent=1, vertical="center")
                cell.font = font_reg
            elif col_name == "2 घंटे में वृद्धि (Plots)":
                cell.fill = fill_positive
                cell.font = font_green
            else:
                cell.font = font_reg
        cur_r += 1

    ws.row_dimensions[cur_r].height = 26
    tot_prev_pct = round((tot_prev / tot_plots * 100), 2) if tot_plots > 0 else 0.0
    tot_curr_pct = round((tot_today / tot_plots * 100), 2) if tot_plots > 0 else 0.0

    t_vals = ["", "महायोग (Grand Total)", tot_plots, tot_prev, tot_today, f"+{tot_growth}", f"{tot_prev_pct:.2f}%",
              f"{tot_curr_pct:.2f}%", f"{(tot_curr_pct - tot_prev_pct):+.2f}%"]
    for c_idx, val in enumerate(t_vals, 1):
        cell = ws.cell(row=cur_r, column=c_idx, value=val)
        cell.font = font_bold
        cell.fill = fill_tot
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for c in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18
    ws.column_dimensions["B"].width = 24

    safe_save_excel(wb, output_excel)

    rows_html = ""
    for _, r in df_result.iterrows():
        g_val = r["2 घंटे में वृद्धि (Plots)"]
        badge_html = f"<span style='background-color:#f3e5f5; color:#6a1b9a; padding:3px 10px; border-radius:4px; font-weight:bold;'>+{g_val}</span>" if g_val > 0 else f"<span style='color:#666;'>{g_val}</span>"
        rows_html += f"""
        <tr>
            <td>{r['स.क्र.']}</td>
            <td style="text-align: left; padding-left: 12px;">{r['तहसील का नाम']}</td>
            <td>{r['कुल Plots']}</td>
            <td>{r['2 घंटे पूर्व पूर्ण']}</td>
            <td>{r['वर्तमान पूर्ण']}</td>
            <td>{badge_html}</td>
            <td>{r['2 घंटे पूर्व %']}</td>
            <td>{r['वर्तमान %']}</td>
            <td style="font-weight:bold; color: #6a1b9a;">{r['वृद्धि %']}</td>
        </tr>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <style>
        @page {{ size: A4 landscape; margin: 8mm 10mm 8mm 10mm; }}
        body {{ font-family: 'Mangal', 'Segoe UI', sans-serif; color: #0F172A; margin: 0; padding: 0; }}
        .header-top {{ display: flex; justify-content: space-between; align-items: center; background: linear-gradient(135deg, #6a1b9a 0%, #ab47bc 100%); color: white; padding: 9px 16px; border-radius: 6px; margin-bottom: 6px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        .header-title {{ font-size: 13pt; font-weight: bold; }}
        .header-dt {{ font-size: 9pt; text-align: right; }}
        .summary-banner {{ background-color: #F3E5F5; color: #4A148C; padding: 7px 12px; font-size: 9.5pt; border-left: 5px solid #8E24AA; border-radius: 4px; margin-bottom: 8px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 9.5pt; text-align: center; }}
        th, td {{ border: 1px solid #CBD5E1; padding: 5px 4px; }}
        th {{ background-color: #374151 !important; color: white !important; font-size: 10pt; font-weight: 600; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        tr:nth-child(even) {{ background-color: #F8FAFC; }}
        .total-row {{ background-color: #E2E8F0 !important; font-weight: bold; font-size: 10pt; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    </style>
    </head>
    <body>
        <div class="header-top">
            <div class="header-title">स्वामित्व योजना - तहसीलवार 2 घंटे पूर्व से तुलनात्मक प्रगति रिपोर्ट | जिला {TARGET_DISTRICT}</div>
            <div class="header-dt">दिनांक: {dt_str}</div>
        </div>
        <div class="summary-banner">
            ⏱️ तुलना आधार: <strong>{base_label}</strong> | 2 घंटे में जिले में कुल नई वृद्धि: <strong>+{tot_growth} Plots</strong>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 6%;">स.क्र.</th>
                    <th style="width: 18%;">तहसील का नाम</th>
                    <th style="width: 11%;">कुल Plots</th>
                    <th style="width: 11%;">2 घंटे पूर्व पूर्ण</th>
                    <th style="width: 11%;">वर्तमान पूर्ण</th>
                    <th style="width: 13%;">2 घंटे में वृद्धि (Plots)</th>
                    <th style="width: 10%;">2 घंटे पूर्व %</th>
                    <th style="width: 10%;">वर्तमान %</th>
                    <th style="width: 10%;">वृद्धि %</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
                <tr class="total-row">
                    <td></td>
                    <td>महायोग</td>
                    <td>{tot_plots}</td>
                    <td>{tot_prev}</td>
                    <td>{tot_today}</td>
                    <td>+{tot_growth}</td>
                    <td>{tot_prev_pct:.2f}%</td>
                    <td>{tot_curr_pct:.2f}%</td>
                    <td style="color:#6a1b9a;">+{(tot_curr_pct - tot_prev_pct):.2f}%</td>
                </tr>
            </tbody>
        </table>
    </body>
    </html>
    """

    temp_html = os.path.join(BASE_DIR, "temp_compare_2hour.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    driver.get(f"file:///{temp_html}")
    time.sleep(1)

    print_options = {
        'landscape': True, 'printBackground': True, 'preferCSSPageSize': True,
        'footerTemplate': '<div style="font-size:9px; text-align:right; width:95%; font-family:Mangal;">पृष्ठ <span class="pageNumber"></span> / <span class="totalPages"></span></div>',
        'headerTemplate': '<div></div>'
    }
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(output_pdf, 'wb') as f:
        f.write(base64.b64decode(result['data']))

    if os.path.exists(temp_html):
        os.remove(temp_html)


# =========================================================
# 11. Tehsilwise Completed Only Comparison Report
# =========================================================
def generate_tehsilwise_completed_comparison_report(driver, goshwara_df, global_tot_col, global_comp_col, output_excel,
                                                    output_pdf):
    today_date = datetime.now().strftime("%Y-%m-%d")
    dt_str = get_current_time_str()
    base_dt = get_morning_baseline_datetime()

    prev_map = {}
    if base_dt:
        conn = sqlite3.connect(DB_PATH)
        prev_df = pd.read_sql(
            f"SELECT tehsil_name, completed FROM tehsil_daily_history WHERE run_datetime = '{base_dt}'", conn)
        conn.close()
        for _, r in prev_df.iterrows():
            prev_map[r['tehsil_name']] = int(r['completed'])

    base_label = f"Morning Baseline ({base_dt})" if base_dt else "Baseline (0)"

    comp_records = []
    tot_prev = 0
    tot_today = 0
    tot_growth = 0

    for _, r in goshwara_df.iterrows():
        t_name = str(r["तहसील"]).strip()
        if "महायोग" in t_name or "Grand Total" in t_name:
            continue

        today_c = int(r.get(global_comp_col, 0))
        prev_c = prev_map.get(t_name, 0)
        diff = today_c - prev_c
        growth_rate = round((diff / prev_c * 100), 2) if prev_c > 0 else (100.0 if diff > 0 else 0.0)

        tot_prev += prev_c
        tot_today += today_c
        tot_growth += diff

        comp_records.append({
            "तहसील का नाम": t_name,
            "पूर्व पूर्ण Plots": prev_c,
            "आज पूर्ण Plots": today_c,
            "शुद्ध वृद्धि (Plots)": diff,
            "वृद्धि दर (%)": growth_rate
        })

    df_result = pd.DataFrame(comp_records)
    df_result = df_result.sort_values(by="शुद्ध वृद्धि (Plots)", ascending=False).reset_index(drop=True)
    df_result.insert(0, "स.क्र.", range(1, len(df_result) + 1))

    wb = Workbook()
    ws = wb.active
    ws.title = "Completed_Comparison"
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    font_title = Font(name="Mangal", size=13, bold=True, color="FFFFFF")
    font_header = Font(name="Mangal", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Mangal", size=11, bold=True)
    font_reg = Font(name="Mangal", size=11)
    font_green = Font(name="Calibri", size=11, bold=True, color="047857")
    font_alert = Font(name="Calibri", size=11, bold=True, color="B91C1C")

    fill_title = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_header = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    fill_zero = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_positive = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    fill_tot = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    ws.merge_cells("A1:F1")
    ws["A1"].value = f"रिपोर्ट समय: {dt_str} | आधार: {base_label}"
    ws["A1"].font = Font(name="Mangal", size=9, italic=True)
    ws["A1"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[1].height = 18

    ws.merge_cells("A2:F2")
    ws["A2"].value = f"स्वामित्व योजना - तहसीलवार केवल Completed Plots तुलनात्मक समीक्षा | जिला {TARGET_DISTRICT}"
    ws["A2"].font = font_title
    ws["A2"].fill = fill_title
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 34

    headers = ["स.क्र.", "तहसील का नाम", "पूर्व पूर्ण Plots", "आज पूर्ण Plots", "शुद्ध वृद्धि (New Plots)",
               "वृद्धि दर (%)"]
    ws.row_dimensions[3].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cur_r = 4
    for _, r in df_result.iterrows():
        ws.row_dimensions[cur_r].height = 22
        growth_val = r["शुद्ध वृद्धि (Plots)"]
        is_zero = (growth_val <= 0)

        c1 = ws.cell(row=cur_r, column=1, value=r["स.क्र."])
        c2 = ws.cell(row=cur_r, column=2, value=r["तहसील का नाम"])
        c3 = ws.cell(row=cur_r, column=3, value=r["पूर्व पूर्ण Plots"])
        c4 = ws.cell(row=cur_r, column=4, value=r["आज पूर्ण Plots"])
        c5 = ws.cell(row=cur_r, column=5, value=f"+{growth_val}" if growth_val > 0 else str(growth_val))
        c6 = ws.cell(row=cur_r, column=6, value=f"{r['वृद्धि दर (%)']:+.2f}%")

        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2.alignment = Alignment(horizontal="left", indent=1, vertical="center")
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c5.alignment = Alignment(horizontal="center", vertical="center")
        c6.alignment = Alignment(horizontal="center", vertical="center")

        for cell in [c1, c2, c3, c4, c5, c6]:
            cell.font = font_reg
            cell.border = thin_border

        if is_zero:
            c5.fill = fill_zero
            c5.font = font_alert
            c6.font = font_alert
        else:
            c5.fill = fill_positive
            c5.font = font_green
            c6.font = font_green

        cur_r += 1

    ws.row_dimensions[cur_r].height = 26
    tot_rate = round((tot_growth / tot_prev * 100), 2) if tot_prev > 0 else 0.0
    tot_vals = ["", "महायोग (Grand Total)", tot_prev, tot_today, f"+{tot_growth}", f"{tot_rate:+.2f}%"]
    for c_idx, val in enumerate(tot_vals, 1):
        cell = ws.cell(row=cur_r, column=c_idx, value=val)
        cell.font = font_bold
        cell.fill = fill_tot
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 24
    ws.column_dimensions["E"].width = 24
    ws.column_dimensions["F"].width = 18

    safe_save_excel(wb, output_excel)

    rows_html = ""
    for _, r in df_result.iterrows():
        g_val = r["शुद्ध वृद्धि (Plots)"]
        if g_val > 0:
            badge_html = f"<span style='background-color:#d1fae5; color:#047857; padding:3px 10px; border-radius:4px; font-weight:bold;'>+{g_val}</span>"
            pct_style = "color:#047857; font-weight:bold;"
        else:
            badge_html = f"<span style='background-color:#fee2e2; color:#b91c1c; padding:3px 10px; border-radius:4px; font-weight:bold;'>0</span>"
            pct_style = "color:#b91c1c; font-weight:bold;"

        rows_html += f"""
        <tr>
            <td>{r['स.क्र.']}</td>
            <td style="text-align: left; padding-left: 14px;">{r['तहसील का नाम']}</td>
            <td>{r['पूर्व पूर्ण Plots']}</td>
            <td>{r['आज पूर्ण Plots']}</td>
            <td>{badge_html}</td>
            <td style="{pct_style}">{r['वृद्धि दर (%)']:+.2f}%</td>
        </tr>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <style>
        @page {{ size: A4 landscape; margin: 10mm 12mm 10mm 12mm; }}
        body {{ font-family: 'Mangal', 'Segoe UI', sans-serif; color: #0F172A; margin: 0; padding: 0; }}
        .header-top {{ display: flex; justify-content: space-between; align-items: center; background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%); color: white; padding: 9px 16px; border-radius: 6px; margin-bottom: 8px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        .header-title {{ font-size: 13.5pt; font-weight: bold; }}
        .header-dt {{ font-size: 9pt; text-align: right; }}
        .summary-banner {{ background-color: #EFF6FF; color: #1E40AF; padding: 8px 14px; font-size: 10pt; border-left: 5px solid #2563EB; border-radius: 4px; margin-bottom: 10px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 10pt; text-align: center; }}
        th, td {{ border: 1px solid #CBD5E1; padding: 6.5px 5px; }}
        th {{ background-color: #1E293B !important; color: white !important; font-size: 10.5pt; font-weight: 600; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        tr:nth-child(even) {{ background-color: #F8FAFC; }}
        .total-row {{ background-color: #E2E8F0 !important; font-weight: bold; font-size: 10.5pt; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    </style>
    </head>
    <body>
        <div class="header-top">
            <div class="header-title">स्वामित्व योजना - तहसीलवार केवल Completed Plots तुलनात्मक समीक्षा</div>
            <div class="header-dt">दिनांक: {today_date}</div>
        </div>
        <div class="summary-banner">
            📊 तुलना आधार: <strong>{base_label}</strong> | कुल शुद्ध रजिस्ट्री वृद्धि: <strong>+{tot_growth} Plots</strong>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 8%;">स.क्र.</th>
                    <th style="width: 24%;">तहसील का नाम</th>
                    <th style="width: 17%;">पूर्व पूर्ण Plots</th>
                    <th style="width: 17%;">आज पूर्ण Plots</th>
                    <th style="width: 18%;">दैनिक वृद्धि (New Plots)</th>
                    <th style="width: 16%;">वृद्धि दर (%)</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
                <tr class="total-row">
                    <td></td>
                    <td>महायोग</td>
                    <td>{tot_prev}</td>
                    <td>{tot_today}</td>
                    <td><span style="color:#047857;">+{tot_growth}</span></td>
                    <td style="color:#047857;">{tot_rate:+.2f}%</td>
                </tr>
            </tbody>
        </table>
    </body>
    </html>
    """

    temp_html = os.path.join(BASE_DIR, "temp_comp_only.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    driver.get(f"file:///{temp_html}")
    time.sleep(1)

    print_options = {
        'landscape': True, 'printBackground': True, 'preferCSSPageSize': True,
        'footerTemplate': '<div style="font-size:9.5px; text-align:right; width:95%; font-family:Mangal;">पृष्ठ <span class="pageNumber"></span> / <span class="totalPages"></span></div>',
        'headerTemplate': '<div></div>'
    }
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(output_pdf, 'wb') as f:
        f.write(base64.b64decode(result['data']))

    if os.path.exists(temp_html):
        os.remove(temp_html)


# =========================================================
# 12. Executive One-Pager Summary Report (With Patwari Names)
# =========================================================
def generate_one_pager_executive_summary(driver, goshwara_df, all_patwaris_df, dist_totals_row, global_pct_col,
                                         global_comp_col, output_excel, output_pdf):
    dt_str = get_current_time_str()
    tot_plots = dist_totals_row.get("कुलप्लॉट", dist_totals_row.get("Total Plots", dist_totals_row.get("कुल Plots", 0)))
    if tot_plots == 0:
        for c in goshwara_df.columns:
            if any(k in str(c).lower() for k in ["plot", "प्लॉट", "किता"]):
                tot_plots = goshwara_df[c].sum()
                break

    tot_beneficiaries = dist_totals_row.get("कुल हितग्राही", dist_totals_row.get("Beneficiaries", 0))
    if tot_beneficiaries == 0:
        for c in goshwara_df.columns:
            if any(k in str(c).lower() for k in ["beneficiar", "हितग्राही"]):
                tot_beneficiaries = goshwara_df[c].sum()
                break

    subm_col = find_column_by_keywords(goshwara_df.columns, ["submitted", "सबमिट", "जमा"])
    tot_submitted = dist_totals_row.get(subm_col, 0) if subm_col else 0
    tot_completed = dist_totals_row.get(global_comp_col, 0) if global_comp_col else 0
    tot_pct = dist_totals_row.get(global_pct_col, 0.0)

    tehsil_sorted = goshwara_df.sort_values(by=global_pct_col, ascending=False).reset_index(drop=True)
    top_3_tehsils = tehsil_sorted.head(3)
    bottom_3_tehsils = tehsil_sorted.tail(3).iloc[::-1]

    # Patwari Name Column Extraction Fix
    patwari_col = None
    for c in all_patwaris_df.columns:
        if any(k in str(c).lower() for k in ["पटवारी", "हल्का", "patwari", "halka"]):
            patwari_col = c
            break
    if not patwari_col:
        patwari_col = all_patwaris_df.columns[1]

    clean_p_df = all_patwaris_df[all_patwaris_df[global_comp_col] != ""].copy()
    patwari_sorted = clean_p_df.sort_values(by=[global_comp_col, global_pct_col], ascending=[False, False]).reset_index(
        drop=True)
    top_5_patwaris = patwari_sorted.head(5)
    bottom_5_patwaris = patwari_sorted.tail(5).iloc[::-1]

    top_tehsil_rows_html = "".join([
                                       f"<tr><td>{idx + 1}</td><td style='text-align:left; padding-left:6px;'>{r['तहसील']}</td><td>{r[global_comp_col]}</td><td>{r[global_pct_col]:.2f}%</td></tr>"
                                       for idx, r in top_3_tehsils.iterrows()])
    bot_tehsil_rows_html = "".join([
                                       f"<tr><td>{len(tehsil_sorted) - idx}</td><td style='text-align:left; padding-left:6px;'>{r['तहसील']}</td><td>{r[global_comp_col]}</td><td>{r[global_pct_col]:.2f}%</td></tr>"
                                       for idx, r in bottom_3_tehsils.reset_index(drop=True).iterrows()])

    # Patwari Rows with Names explicitly resolved
    top_patwari_rows_html = "".join([
                                        f"<tr><td>{idx + 1}</td><td style='text-align:left; padding-left:6px;'>{r['तहसील']} - {r[patwari_col]}</td><td>{r[global_comp_col]}</td><td>{float(r[global_pct_col]):.2f}%</td></tr>"
                                        for idx, r in top_5_patwaris.iterrows()])
    bot_patwari_rows_html = "".join([
                                        f"<tr><td>{len(patwari_sorted) - idx}</td><td style='text-align:left; padding-left:6px;'>{r['तहसील']} - {r[patwari_col]}</td><td>{r[global_comp_col]}</td><td>{float(r[global_pct_col]):.2f}%</td></tr>"
                                        for idx, r in bottom_5_patwaris.reset_index(drop=True).iterrows()])

    html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <style>
        @page {{ size: A4 landscape; margin: 8mm 8mm 8mm 8mm; }}
        body {{ font-family: 'Mangal', 'Nirmala UI', sans-serif; color: #111; margin: 0; padding: 0; font-size: 10px; }}
        .header-top {{ display: flex; justify-content: space-between; align-items: center; background-color: #1F497D; color: white; padding: 8px 14px; border-radius: 4px; margin-bottom: 8px; }}
        .header-title {{ font-size: 14pt; font-weight: bold; }}
        .header-dt {{ font-size: 9pt; text-align: right; }}
        .kpi-container {{ display: flex; justify-content: space-between; gap: 8px; margin-bottom: 10px; }}
        .kpi-card {{ flex: 1; background-color: #f2f5f9; border: 1.5px solid #b8cce4; border-radius: 5px; text-align: center; padding: 6px 4px; }}
        .kpi-label {{ font-size: 8.5pt; font-weight: bold; color: #333; margin-bottom: 4px; }}
        .kpi-value {{ font-size: 14pt; font-weight: bold; color: #0d47a1; font-family: 'Calibri', sans-serif; }}
        .two-col-grid {{ display: flex; gap: 12px; margin-bottom: 10px; }}
        .col-half {{ flex: 1; }}
        .section-header-green {{ background-color: #385723; color: white; padding: 5px 8px; font-weight: bold; font-size: 9.5pt; border-radius: 4px 4px 0 0; text-align: center; }}
        .section-header-red {{ background-color: #c00000; color: white; padding: 5px 8px; font-weight: bold; font-size: 9.5pt; border-radius: 4px 4px 0 0; text-align: center; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 9pt; text-align: center; }}
        th, td {{ border: 1px solid #777; padding: 4.5px 4px; }}
        th {{ background-color: #41719C; color: white; font-weight: bold; }}
        .row-green {{ background-color: #e2efda; }}
        .row-red {{ background-color: #fce4d6; }}
    </style>
    </head>
    <body>
        <div class="header-top">
            <div class="header-title">स्वामित्व योजना एक-दृष्टि प्रगति सारांश (Executive One-Pager) - जिला {TARGET_DISTRICT}</div>
            <div class="header-dt">दिनांक व समय: {dt_str}</div>
        </div>
        <div class="kpi-container">
            <div class="kpi-card"><div class="kpi-label">1. Total Target (Plots)</div><div class="kpi-value">{tot_plots}</div></div>
            <div class="kpi-card"><div class="kpi-label">2. No. of Beneficiaries</div><div class="kpi-value">{tot_beneficiaries}</div></div>
            <div class="kpi-card"><div class="kpi-label">3. Submitted (Sampda)</div><div class="kpi-value">{tot_submitted}</div></div>
            <div class="kpi-card"><div class="kpi-label">4. Completed (Registry)</div><div class="kpi-value">{tot_completed}</div></div>
            <div class="kpi-card"><div class="kpi-label">5. Overall Percentage</div><div class="kpi-value">{tot_pct:.2f}%</div></div>
        </div>
        <div class="two-col-grid">
            <div class="col-half">
                <div class="section-header-green">6. शीर्ष प्रदर्शनकारी तहसीलें (Top 3 Tehsils)</div>
                <table><thead><tr><th style="width: 12%;">रैंक</th><th>तहसील</th><th style="width: 24%;">पूर्ण Plots</th><th style="width: 24%;">प्रगति (%)</th></tr></thead><tbody class="row-green">{top_tehsil_rows_html}</tbody></table>
            </div>
            <div class="col-half">
                <div class="section-header-red">7. न्यूनतम प्रगतिशील तहसीलें (Bottom 3 Tehsils)</div>
                <table><thead><tr><th style="width: 12%;">रैंक</th><th>तहसील</th><th style="width: 24%;">पूर्ण Plots</th><th style="width: 24%;">प्रगति (%)</th></tr></thead><tbody class="row-red">{bot_tehsil_rows_html}</tbody></table>
            </div>
        </div>
        <div class="two-col-grid">
            <div class="col-half">
                <div class="section-header-green">8. शीर्ष 5 पटवारी (Top 5 Performing Patwaris)</div>
                <table><thead><tr><th style="width: 12%;">रैंक</th><th>तहसील / हल्का (पटवारी)</th><th style="width: 24%;">पूर्ण Plots</th><th style="width: 24%;">प्रगति (%)</th></tr></thead><tbody class="row-green">{top_patwari_rows_html}</tbody></table>
            </div>
            <div class="col-half">
                <div class="section-header-red">9. न्यूनतम 5 पटवारी (Bottom 5 Patwaris)</div>
                <table><thead><tr><th style="width: 12%;">रैंक</th><th>तहसील / हल्का (पटवारी)</th><th style="width: 24%;">पूर्ण Plots</th><th style="width: 24%;">प्रगति (%)</th></tr></thead><tbody class="row-red">{bot_patwari_rows_html}</tbody></table>
            </div>
        </div>
    </body>
    </html>
    """
    temp_html = os.path.join(BASE_DIR, "temp_one_pager.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    driver.get(f"file:///{temp_html}")
    time.sleep(1)

    print_options = {'landscape': True, 'printBackground': True, 'preferCSSPageSize': True}
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(output_pdf, 'wb') as f:
        f.write(base64.b64decode(result['data']))

    if os.path.exists(temp_html):
        os.remove(temp_html)


# =========================================================
# 13. Tehsilwise Bottom 10 Patwaris (Separate Pages)
# =========================================================
def generate_tehsilwise_bottom_10_patwaris_report(driver, all_tehsils_patwari_list, tot_plot_col, comp_col,
                                                  output_excel, output_pdf):
    dt_str = get_current_time_str()
    wb = Workbook()
    ws_default = wb.active
    wb.remove(ws_default)

    font_title = Font(name="Mangal", size=13, bold=True, color="FFFFFF")
    font_header = Font(name="Mangal", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Mangal", size=11, bold=True)
    font_reg = Font(name="Mangal", size=11)
    font_dt = Font(name="Mangal", size=10, italic=True)

    fill_title = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_header = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    fill_zero = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_tot = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='B0B0B0'), right=Side(style='thin', color='B0B0B0'),
        top=Side(style='thin', color='B0B0B0'), bottom=Side(style='thin', color='B0B0B0')
    )

    tehsils_pages_html = ""

    for df_t in all_tehsils_patwari_list:
        if df_t.empty:
            continue

        t_name = str(df_t["तहसील"].iloc[0]).strip()
        patwari_col = df_t.columns[1]

        df_filtered = df_t[df_t[tot_plot_col] != df_t[comp_col]].copy()
        if df_filtered.empty:
            continue

        sorted_t = df_filtered.sort_values(by=[comp_col, tot_plot_col], ascending=[True, True]).head(10).copy()

        clean_sheet_name = t_name[:30].replace("/", "_").replace("\\", "_")
        ws = wb.create_sheet(title=clean_sheet_name)
        ws.views.sheetView[0].showGridLines = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4

        ws.merge_cells("A1:E1")
        ws["A1"].value = f"रिपोर्ट दिनांक व समय: {dt_str}"
        ws["A1"].font = font_dt
        ws["A1"].alignment = Alignment(horizontal="right", vertical="center")

        ws.merge_cells("A2:E2")
        ws["A2"].value = f"स्वामित्व योजना - तहसील: {t_name} के न्यूनतम 10 पटवारी (Bottom 10 by Completion)"
        ws["A2"].font = font_title
        ws["A2"].fill = fill_title
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

        headers = ["स.क्र.", "तहसील", "पटवारी हल्का", "No. of Plots", "Completed"]
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        cur_r = 4
        tot_p = 0
        tot_c = 0
        rows_html_part = ""

        for s_no, (_, r) in enumerate(sorted_t.iterrows(), start=1):
            p_val = int(r[tot_plot_col])
            c_val = int(r[comp_col])
            tot_p += p_val
            tot_c += c_val

            c1 = ws.cell(row=cur_r, column=1, value=s_no)
            c2 = ws.cell(row=cur_r, column=2, value=t_name)
            c3 = ws.cell(row=cur_r, column=3, value=r[patwari_col])
            c4 = ws.cell(row=cur_r, column=4, value=p_val)
            c5 = ws.cell(row=cur_r, column=5, value=c_val)

            c1.alignment = Alignment(horizontal="center", vertical="center")
            c2.alignment = Alignment(horizontal="center", vertical="center")
            c3.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            c4.alignment = Alignment(horizontal="center", vertical="center")
            c5.alignment = Alignment(horizontal="center", vertical="center")

            for cell in [c1, c2, c3, c4, c5]:
                cell.font = font_reg
                cell.border = thin_border
                if c_val == 0:
                    cell.font = font_bold
                    cell.fill = fill_zero

            tr_bg = "background-color: #fce4d6; font-weight: bold;" if c_val == 0 else ""
            rows_html_part += f"""
            <tr style="{tr_bg}">
                <td>{s_no}</td>
                <td>{t_name}</td>
                <td style="text-align: left; padding-left: 12px;">{r[patwari_col]}</td>
                <td>{p_val}</td>
                <td>{c_val}</td>
            </tr>
            """
            cur_r += 1

        ws.cell(row=cur_r, column=1, value="")
        c_tot_lbl = ws.cell(row=cur_r, column=2, value="योग (Total)")
        c_tot_cnt = ws.cell(row=cur_r, column=3, value=f"{len(sorted_t)} पटवारी")
        c_tot_p = ws.cell(row=cur_r, column=4, value=tot_p)
        c_tot_c = ws.cell(row=cur_r, column=5, value=tot_c)

        for c in [ws.cell(row=cur_r, column=1), c_tot_lbl, c_tot_cnt, c_tot_p, c_tot_c]:
            c.font = font_bold
            c.fill = fill_tot
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")

        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 24
        ws.column_dimensions["C"].width = 36
        ws.column_dimensions["D"].width = 22
        ws.column_dimensions["E"].width = 22

        tehsils_pages_html += f"""
        <div class="tehsil-page">
            <div class="header-top">
                <div class="header-title">स्वामित्व योजना - तहसील: {t_name} के न्यूनतम 10 पटवारी (Bottom 10 by Completion)</div>
                <div class="header-dt">दिनांक व समय: {dt_str}</div>
            </div>
            <div class="note-box">
                * नोट: जिन पटवारियों के No. of Plots एवं Completed समान (100% लक्ष्य पूर्ण) हैं, उन्हें इस सूची से पृथक रखा गया है।
            </div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 10%;">स.क्र.</th>
                        <th style="width: 22%;">तहसील</th>
                        <th style="width: 36%;">पटवारी हल्का</th>
                        <th style="width: 16%;">No. of Plots</th>
                        <th style="width: 16%;">Completed</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html_part}
                    <tr class="total-row">
                        <td></td>
                        <td>योग</td>
                        <td style="text-align: left; padding-left: 12px;">{len(sorted_t)} पटवारी</td>
                        <td>{tot_p}</td>
                        <td>{tot_c}</td>
                    </tr>
                </tbody>
            </table>
        </div>
        """

    safe_save_excel(wb, output_excel)

    full_html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <style>
        @page {{ size: A4 landscape; margin: 10mm 12mm 10mm 12mm; }}
        body {{ font-family: 'Mangal', 'Nirmala UI', sans-serif; color: #111; margin: 0; padding: 0; }}
        .tehsil-page {{ page-break-after: always; break-after: page; }}
        .tehsil-page:last-child {{ page-break-after: avoid; break-after: avoid; }}
        .header-top {{ display: flex; justify-content: space-between; align-items: center; background-color: #1F497D; color: white; padding: 8px 14px; border-radius: 4px; margin-bottom: 6px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        .header-title {{ font-size: 13.5pt; font-weight: bold; }}
        .header-dt {{ font-size: 9.5pt; text-align: right; }}
        .note-box {{ font-size: 9.5pt; color: #555; font-style: italic; margin-bottom: 8px; padding: 2px 4px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 11pt; text-align: center; }}
        th, td {{ border: 1px solid #777; padding: 6.5px 5px; }}
        th {{ background-color: #C00000 !important; color: white !important; font-size: 11.5pt; font-weight: bold; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        tr:nth-child(even) {{ background-color: #fafafa; }}
        .total-row {{ background-color: #d9e1f2 !important; font-weight: bold; font-size: 11.5pt; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    </style>
    </head>
    <body>{tehsils_pages_html}</body>
    </html>
    """

    temp_html = os.path.join(BASE_DIR, "temp_bottom_10_pages.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(full_html_content)

    driver.get(f"file:///{temp_html}")
    time.sleep(1)

    print_options = {
        'landscape': True, 'printBackground': True, 'preferCSSPageSize': True,
        'footerTemplate': '<div style="font-size:9.5px; text-align:right; width:95%; font-family:Mangal;">पृष्ठ <span class="pageNumber"></span> / <span class="totalPages"></span></div>',
        'headerTemplate': '<div></div>'
    }
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(output_pdf, 'wb') as f:
        f.write(base64.b64decode(result['data']))

    if os.path.exists(temp_html):
        os.remove(temp_html)


# =========================================================
# 14. All Patwaris Report (<50 Highlight + Separate Pages)
# =========================================================
def generate_all_patwaris_tehsilwise_custom_report(driver, all_tehsils_patwari_list, tot_plot_col, comp_col,
                                                   pct_col_name, output_excel, output_pdf):
    dt_str = get_current_time_str()
    wb = Workbook()
    ws_default = wb.active
    wb.remove(ws_default)

    font_title = Font(name="Mangal", size=13, bold=True, color="FFFFFF")
    font_header = Font(name="Mangal", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Mangal", size=10, bold=True)
    font_reg = Font(name="Mangal", size=10)
    font_warn = Font(name="Mangal", size=10, bold=True, color="9F1239")
    font_dt = Font(name="Mangal", size=9, italic=True)

    fill_title = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_header = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
    fill_warn_card = PatternFill(start_color="FFE4E6", end_color="FFE4E6", fill_type="solid")
    fill_under_50 = PatternFill(start_color="FFF1F2", end_color="FFF1F2", fill_type="solid")
    fill_tot = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    tehsils_pages_html = ""

    for df_t in all_tehsils_patwari_list:
        if df_t.empty:
            continue

        t_name = str(df_t["तहसील"].iloc[0]).strip()
        patwari_col = df_t.columns[1]
        less_50_count = int((df_t[comp_col] < 50).sum())

        clean_sheet_name = t_name[:30].replace("/", "_").replace("\\", "_")
        ws = wb.create_sheet(title=clean_sheet_name)
        ws.views.sheetView[0].showGridLines = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4

        ws.merge_cells("A1:G1")
        ws["A1"].value = f"रिपोर्ट दिनांक व समय: {dt_str}"
        ws["A1"].font = font_dt
        ws["A1"].alignment = Alignment(horizontal="right", vertical="center")

        ws.merge_cells("A2:G2")
        ws["A2"].value = f"स्वामित्व योजना पटवारीवार प्रगति - तहसील: {t_name} | जिला {TARGET_DISTRICT}"
        ws["A2"].font = font_title
        ws["A2"].fill = fill_title
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 32

        ws.merge_cells("A3:G3")
        ws["A3"].value = f"⚠️ विशेष निगरानी: 50 से कम पूर्ण (Completed < 50) वाले कुल पटवारी: {less_50_count}"
        ws["A3"].font = font_warn
        ws["A3"].fill = fill_warn_card
        ws["A3"].alignment = Alignment(horizontal="left", indent=1, vertical="center")
        ws.row_dimensions[3].height = 24

        headers = ["स.क्र.", "तहसील", "पटवारी हल्का", "कुल ग्राम", "No. of Plots", "Completed", "प्रगति (%)"]
        ws.row_dimensions[4].height = 28
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        cur_r = 5
        rows_html_part = ""

        for s_no, (_, r) in enumerate(df_t.iterrows(), start=1):
            c_val = int(r[comp_col])
            is_under_50 = (c_val < 50)

            c1 = ws.cell(row=cur_r, column=1, value=s_no)
            c2 = ws.cell(row=cur_r, column=2, value=t_name)
            c3 = ws.cell(row=cur_r, column=3, value=r[patwari_col])
            c4 = ws.cell(row=cur_r, column=4, value=int(r["कुल ग्राम"]))
            c5 = ws.cell(row=cur_r, column=5, value=int(r[tot_plot_col]))
            c6 = ws.cell(row=cur_r, column=6, value=c_val)
            c7 = ws.cell(row=cur_r, column=7, value=f"{float(r[pct_col_name]):.2f}%")

            c1.alignment = Alignment(horizontal="center", vertical="center")
            c2.alignment = Alignment(horizontal="center", vertical="center")
            c3.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            c4.alignment = Alignment(horizontal="center", vertical="center")
            c5.alignment = Alignment(horizontal="center", vertical="center")
            c6.alignment = Alignment(horizontal="center", vertical="center")
            c7.alignment = Alignment(horizontal="center", vertical="center")

            for cell in [c1, c2, c3, c4, c5, c6, c7]:
                cell.border = thin_border
                if is_under_50:
                    cell.font = font_warn
                    cell.fill = fill_under_50
                else:
                    cell.font = font_reg

            row_class = "row-under-50" if is_under_50 else ""
            badge_html = f'<span class="badge-alert">{c_val}</span>' if is_under_50 else str(c_val)

            rows_html_part += f"""
            <tr class="{row_class}">
                <td>{s_no}</td>
                <td>{t_name}</td>
                <td style="text-align: left; padding-left: 12px; font-weight: {'600' if is_under_50 else 'normal'};">{r[patwari_col]}</td>
                <td>{int(r['कुल ग्राम'])}</td>
                <td>{int(r[tot_plot_col])}</td>
                <td>{badge_html}</td>
                <td style="font-weight: 600;">{float(r[pct_col_name]):.2f}%</td>
            </tr>
            """
            cur_r += 1

        t_plots = int(df_t[tot_plot_col].sum())
        t_comp = int(df_t[comp_col].sum())
        t_pct = round((t_comp / t_plots * 100), 2) if t_plots > 0 else 0.0

        ws.row_dimensions[cur_r].height = 26
        ws.cell(row=cur_r, column=1, value="")
        c_tot_lbl = ws.cell(row=cur_r, column=2, value="तहसील महायोग")
        ws.cell(row=cur_r, column=3, value=f"{len(df_t)} पटवारी")
        ws.cell(row=cur_r, column=4, value=int(df_t["कुल ग्राम"].sum()))
        ws.cell(row=cur_r, column=5, value=t_plots)
        ws.cell(row=cur_r, column=6, value=t_comp)
        ws.cell(row=cur_r, column=7, value=f"{t_pct:.2f}%")

        for c in [ws.cell(row=cur_r, column=i) for i in range(1, 8)]:
            c.font = font_bold
            c.fill = fill_tot
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")

        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 32
        ws.column_dimensions["D"].width = 16
        ws.column_dimensions["E"].width = 18
        ws.column_dimensions["F"].width = 18
        ws.column_dimensions["G"].width = 18

        tehsils_pages_html += f"""
        <div class="tehsil-page">
            <div class="header-container">
                <div class="header-title">स्वामित्व योजना पटवारी प्रगति रिपोर्ट - तहसील: {t_name}</div>
                <div class="header-dt">दिनांक व समय: {dt_str}</div>
            </div>
            <div class="alert-banner">
                <span class="alert-icon">⚠️</span> 
                <span>विशेष निगरानी: <strong>50 से कम पूर्ण (Completed &lt; 50)</strong> वाले कुल पटवारी: <strong>{less_50_count}</strong></span>
            </div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 7%;">स.क्र.</th>
                        <th style="width: 17%;">तहसील</th>
                        <th style="width: 32%;">पटवारी हल्का</th>
                        <th style="width: 11%;">कुल ग्राम</th>
                        <th style="width: 11%;">No. of Plots</th>
                        <th style="width: 11%;">Completed</th>
                        <th style="width: 11%;">प्रगति (%)</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html_part}
                    <tr class="total-row">
                        <td></td>
                        <td>तहसील महायोग</td>
                        <td style="text-align: left; padding-left: 12px;">{len(df_t)} पटवारी</td>
                        <td>{int(df_t['कुल ग्राम'].sum())}</td>
                        <td>{t_plots}</td>
                        <td>{t_comp}</td>
                        <td>{t_pct:.2f}%</td>
                    </tr>
                </tbody>
            </table>
        </div>
        """

    safe_save_excel(wb, output_excel)

    full_html_content = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <style>
        @page {{ size: A4 landscape; margin: 10mm 12mm 10mm 12mm; }}
        body {{ font-family: 'Mangal', 'Segoe UI', sans-serif; color: #0F172A; margin: 0; padding: 0; background-color: #FFFFFF; }}
        .tehsil-page {{ page-break-after: always; break-after: page; }}
        .tehsil-page:last-child {{ page-break-after: avoid; break-after: avoid; }}
        .header-container {{ display: flex; justify-content: space-between; align-items: center; background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%); color: #FFFFFF; padding: 9px 16px; border-radius: 6px; margin-bottom: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.06); -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        .header-title {{ font-size: 13.5pt; font-weight: 700; letter-spacing: 0.3px; }}
        .header-dt {{ font-size: 9pt; opacity: 0.95; font-style: italic; }}
        .alert-banner {{ background-color: #FFF1F2; color: #9F1239; padding: 7px 14px; font-size: 9.5pt; border-left: 5px solid #E11D48; border-radius: 4px; margin-bottom: 10px; display: flex; align-items: center; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        .alert-icon {{ font-size: 11pt; margin-right: 8px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 9.5pt; text-align: center; }}
        th, td {{ border: 1px solid #E2E8F0; padding: 5.5px 5px; }}
        th {{ background-color: #1E293B !important; color: #FFFFFF !important; font-size: 10pt; font-weight: 600; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        tr:nth-child(even) {{ background-color: #F8FAFC; }}
        .row-under-50 {{ background-color: #FFF1F2 !important; color: #9F1239; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        .badge-alert {{ background-color: #FFE4E6; color: #BE123C; padding: 2px 8px; border-radius: 4px; font-weight: 700; display: inline-block; border: 1px solid #FECDD3; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        .total-row {{ background-color: #E2E8F0 !important; color: #0F172A; font-weight: 700; font-size: 10pt; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    </style>
    </head>
    <body>{tehsils_pages_html}</body>
    </html>
    """

    temp_html = os.path.join(BASE_DIR, "temp_all_patwaris_tehsil_pages.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(full_html_content)

    driver.get(f"file:///{temp_html}")
    time.sleep(1)

    print_options = {
        'landscape': True, 'printBackground': True, 'preferCSSPageSize': True,
        'footerTemplate': '<div style="font-size:9px; text-align:right; width:95%; font-family:Mangal;">पृष्ठ <span class="pageNumber"></span> / <span class="totalPages"></span></div>',
        'headerTemplate': '<div></div>'
    }
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    with open(output_pdf, 'wb') as f:
        f.write(base64.b64decode(result['data']))

    if os.path.exists(temp_html):
        os.remove(temp_html)


# =========================================================
# 15. PowerPoint Presentation Module
# =========================================================
def generate_review_pptx(pptx_path, df_state_comp, df_state_pct, goshwara_df, all_tehsil_patwaris_list, global_comp_col,
                         global_tot_col, global_pct_col):
    if not PPTX_AVAILABLE:
        logging.warning("python-pptx उपलब्ध नहीं है। PPT स्किप की जा रही है।")
        return

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    c_blue = RGBColor(31, 73, 125)
    c_white = RGBColor(255, 255, 255)
    c_red = RGBColor(192, 0, 0)
    c_gray = RGBColor(242, 245, 249)

    def add_slide_header(slide, title_text):
        top_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.4), Inches(12.333), Inches(0.8))
        top_box.fill.solid()
        top_box.fill.fore_color.rgb = c_blue
        top_box.line.color.rgb = c_blue
        tf = top_box.text_frame
        tf.text = title_text
        p = tf.paragraphs[0]
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = c_white
        p.alignment = PP_ALIGN.CENTER

    # Slide 1: जिला स्तर (55 जिलों में अशोकनगर की रैंक)
    s1 = prs.slides.add_slide(blank_layout)
    add_slide_header(s1, f"स्वामित्व योजना राज्य स्तरीय समीक्षा - जिला: {TARGET_DISTRICT}")

    comp_rank = "N/A"
    pct_rank = "N/A"
    dist_col = df_state_comp.columns[1] if df_state_comp is not None else ""

    if df_state_comp is not None:
        m_c = df_state_comp[df_state_comp[dist_col].astype(str).str.contains(TARGET_DISTRICT, na=False)]
        if not m_c.empty:
            comp_rank = m_c.index[0] + 1

    if df_state_pct is not None:
        m_p = df_state_pct[df_state_pct[dist_col].astype(str).str.contains(TARGET_DISTRICT, na=False)]
        if not m_p.empty:
            pct_rank = m_p.index[0] + 1

    card1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(2.2), Inches(4.8), Inches(3.2))
    card1.fill.solid()
    card1.fill.fore_color.rgb = c_gray
    tf1 = card1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "कम्प्लीशन (पूर्ण संख्या) अनुसार रैंक\n"
    p1.font.size = Pt(18)
    p1.font.bold = True
    p1.alignment = PP_ALIGN.CENTER

    p1_sub = tf1.add_paragraph()
    p1_sub.text = f"{comp_rank} / 55"
    p1_sub.font.size = Pt(44)
    p1_sub.font.bold = True
    p1_sub.font.color.rgb = c_blue
    p1_sub.alignment = PP_ALIGN.CENTER

    card2 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.0), Inches(2.2), Inches(4.8), Inches(3.2))
    card2.fill.solid()
    card2.fill.fore_color.rgb = c_gray
    tf2 = card2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "प्रगति प्रतिशत (%) अनुसार रैंक\n"
    p2.font.size = Pt(18)
    p2.font.bold = True
    p2.alignment = PP_ALIGN.CENTER

    p2_sub = tf2.add_paragraph()
    p2_sub.text = f"{pct_rank} / 55"
    p2_sub.font.size = Pt(44)
    p2_sub.font.bold = True
    p2_sub.font.color.rgb = c_red
    p2_sub.alignment = PP_ALIGN.CENTER

    dt_tx = s1.shapes.add_textbox(Inches(1.5), Inches(6.0), Inches(10.3), Inches(0.5))
    dt_tx.text_frame.text = f"समीक्षा दिनांक व समय: {get_current_time_str()} | स्रोत: MPIGR स्वामित्व डैशबोर्ड"
    dt_tx.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Slide 2: तहसीलवार गोशवारा
    if goshwara_df is not None and not goshwara_df.empty:
        s2 = prs.slides.add_slide(blank_layout)
        add_slide_header(s2, f"तहसीलवार प्रगति गोशवारा - जिला {TARGET_DISTRICT}")

        cols_to_show = ["क्र.", "तहसील", "कुल पटवारी", "कुल ग्राम"]
        other_c = [c for c in goshwara_df.columns if c not in cols_to_show + [global_pct_col]]
        final_cols = cols_to_show + other_c + [global_pct_col]
        final_cols = [c for c in final_cols if c in goshwara_df.columns]

        rows_cnt = min(len(goshwara_df) + 1, 12)
        cols_cnt = len(final_cols)
        table_shape = s2.shapes.add_table(rows_cnt, cols_cnt, Inches(0.5), Inches(1.5), Inches(12.333), Inches(5.2))
        tbl = table_shape.table

        for c_idx, c_name in enumerate(final_cols):
            cell = tbl.cell(0, c_idx)
            cell.text = str(c_name)
            cell.fill.solid()
            cell.fill.fore_color.rgb = c_blue
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(11)
                p.font.bold = True
                p.font.color.rgb = c_white
                p.alignment = PP_ALIGN.CENTER

        for r_idx in range(rows_cnt - 1):
            row_data = goshwara_df.iloc[r_idx]
            for c_idx, c_name in enumerate(final_cols):
                cell = tbl.cell(r_idx + 1, c_idx)
                val = row_data[c_name]
                display_v = f"{float(val):.2f}%" if c_name == global_pct_col else str(val)
                cell.text = display_v
                for p in cell.text_frame.paragraphs:
                    p.font.size = Pt(10)
                    p.alignment = PP_ALIGN.CENTER

    # Slide 3: State Completion
    if df_state_comp is not None and not df_state_comp.empty:
        s3 = prs.slides.add_slide(blank_layout)
        add_slide_header(s3, "राज्य स्तरीय रैंकिंग - Completion (पूर्ण संख्या) अनुसार")

        disp_comp = df_state_comp.head(15).copy()
        t_cols = [c for c in disp_comp.columns[:6]]
        table_s = s3.shapes.add_table(len(disp_comp) + 1, len(t_cols), Inches(0.5), Inches(1.4), Inches(12.333),
                                      Inches(5.5))
        tbl = table_s.table

        for c_idx, c_name in enumerate(t_cols):
            cell = tbl.cell(0, c_idx)
            cell.text = str(c_name)
            cell.fill.solid()
            cell.fill.fore_color.rgb = c_blue
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(10.5)
                p.font.bold = True
                p.font.color.rgb = c_white

        for r_idx in range(len(disp_comp)):
            r_data = disp_comp.iloc[r_idx]
            is_ashok = (TARGET_DISTRICT in str(r_data[dist_col]))
            for c_idx, c_name in enumerate(t_cols):
                cell = tbl.cell(r_idx + 1, c_idx)
                cell.text = str(r_data[c_name])
                if is_ashok:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(255, 255, 0)
                for p in cell.text_frame.paragraphs:
                    p.font.size = Pt(9.5)
                    p.alignment = PP_ALIGN.CENTER

    # Slide 4: State Percentage
    if df_state_pct is not None and not df_state_pct.empty:
        s4 = prs.slides.add_slide(blank_layout)
        add_slide_header(s4, "राज्य स्तरीय रैंकिंग - Percentage (प्रतिशत) अनुसार")

        disp_pct = df_state_pct.head(15).copy()
        t_cols = [c for c in disp_pct.columns[:6]]
        table_s = s4.shapes.add_table(len(disp_pct) + 1, len(t_cols), Inches(0.5), Inches(1.4), Inches(12.333),
                                      Inches(5.5))
        tbl = table_s.table

        for c_idx, c_name in enumerate(t_cols):
            cell = tbl.cell(0, c_idx)
            cell.text = str(c_name)
            cell.fill.solid()
            cell.fill.fore_color.rgb = c_blue
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(10.5)
                p.font.bold = True
                p.font.color.rgb = c_white

        for r_idx in range(len(disp_pct)):
            r_data = disp_pct.iloc[r_idx]
            is_ashok = (TARGET_DISTRICT in str(r_data[dist_col]))
            for c_idx, c_name in enumerate(t_cols):
                cell = tbl.cell(r_idx + 1, c_idx)
                cell.text = str(r_data[c_name])
                if is_ashok:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(255, 255, 0)
                for p in cell.text_frame.paragraphs:
                    p.font.size = Pt(9.5)
                    p.alignment = PP_ALIGN.CENTER

    # Slide 5 onwards: Tehsilwise Bottom 10
    for df_t in all_tehsil_patwaris_list:
        if df_t.empty:
            continue
        t_name = str(df_t["तहसील"].iloc[0]).strip()
        patwari_col = df_t.columns[1]

        df_filtered = df_t[df_t[global_tot_col] != df_t[global_comp_col]].copy()
        if df_filtered.empty:
            continue

        sorted_t = df_filtered.sort_values(by=[global_comp_col, global_tot_col], ascending=[True, True]).head(10).copy()

        st_slide = prs.slides.add_slide(blank_layout)
        add_slide_header(st_slide, f"तहसील: {t_name} - न्यूनतम 10 पटवारी (Bottom 10 by Completion)")

        t_shape = st_slide.shapes.add_table(len(sorted_t) + 1, 5, Inches(1.0), Inches(1.5), Inches(11.333), Inches(5.0))
        tb = t_shape.table

        b_headers = ["स.क्र.", "तहसील", "पटवारी हल्का", "No. of Plots", "Completed"]
        for c_idx, h in enumerate(b_headers):
            c = tb.cell(0, c_idx)
            c.text = h
            c.fill.solid()
            c.fill.fore_color.rgb = c_red
            for p in c.text_frame.paragraphs:
                p.font.size = Pt(11)
                p.font.bold = True
                p.font.color.rgb = c_white
                p.alignment = PP_ALIGN.CENTER

        for r_idx, (_, row) in enumerate(sorted_t.iterrows()):
            vals = [str(r_idx + 1), t_name, str(row[patwari_col]), str(row[global_tot_col]), str(row[global_comp_col])]
            is_zero = (int(row[global_comp_col]) == 0)
            for c_idx, v in enumerate(vals):
                c = tb.cell(r_idx + 1, c_idx)
                c.text = v
                if is_zero:
                    c.fill.solid()
                    c.fill.fore_color.rgb = RGBColor(252, 228, 214)
                for p in c.text_frame.paragraphs:
                    p.font.size = Pt(10)
                    p.alignment = PP_ALIGN.CENTER

    prs.save(pptx_path)
    logging.info(f"-> PowerPoint Presentation (.pptx) तैयार: {pptx_path}")


# =========================================================
# 16. WhatsApp Delivery (Fault-Tolerant & Crash Safe)
# =========================================================
def send_whatsapp_report(target_name_or_number, message_text, file_paths=None, is_group=False):
    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={WHATSAPP_SESSION_DIR}")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    wa_driver = None
    try:
        wa_driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        wait = WebDriverWait(wa_driver, 60)

        if is_group:
            wa_driver.get("https://web.whatsapp.com/")
            logging.info("WhatsApp Web ग्रुप के लिए लोड हो रहा है...")
            time.sleep(6)
            search_box = wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")))
            search_box.clear()
            search_box.send_keys(target_name_or_number)
            time.sleep(2)
            search_box.send_keys(Keys.ENTER)
            time.sleep(3)
        else:
            encoded_msg = quote(message_text)
            url = f"https://web.whatsapp.com/send?phone={target_name_or_number}&text={encoded_msg}"
            wa_driver.get(url)
            logging.info(f"नंबर {target_name_or_number} पर WhatsApp खोला जा रहा है...")

        if not wa_driver.window_handles:
            raise Exception("Chrome Window बंद हो गई।")

        chat_box = wait.until(EC.presence_of_element_located((
            By.XPATH, "//footer//div[@contenteditable='true'] | //div[@data-tab='10'][@contenteditable='true']"
        )))
        time.sleep(2)

        if is_group and message_text:
            chat_box.send_keys(message_text)
            time.sleep(1)
            chat_box.send_keys(Keys.ENTER)
            time.sleep(2)
        elif not is_group:
            chat_box.send_keys(Keys.ENTER)
            time.sleep(2)

        if file_paths:
            for f_path in file_paths:
                if not wa_driver.window_handles:
                    raise Exception("फ़ाइल भेजते समय ब्राउज़र विंडो बंद हो गई।")

                if not os.path.exists(f_path):
                    continue

                attach_btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    "//footer//*[@title='Attach' or @title='संलग्न करें' or @data-icon='plus'] | //div[@title='Attach']"
                )))
                attach_btn.click()
                time.sleep(2)

                doc_input = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[@accept='*' and @type='file']")))
                doc_input.send_keys(os.path.abspath(f_path))
                time.sleep(3)

                try:
                    send_btn = wait.until(EC.element_to_be_clickable((By.XPATH,
                                                                      "//span[@data-icon='send']/parent::button | //div[@aria-label='Send' or @aria-label='भेजें']")))
                    send_btn.click()
                except Exception:
                    caption_box = wa_driver.find_element(By.XPATH, "//div[@contenteditable='true']")
                    caption_box.send_keys(Keys.ENTER)

                logging.info(f"WhatsApp पर फ़ाइल भेजी गई: {os.path.basename(f_path)}")
                time.sleep(4)

        time.sleep(3)
        print("  ✓ WhatsApp संदेश एवं सभी रिपोर्ट्स सफलतापूर्वक डिलीवर हो गईं।")

    except Exception as e:
        logging.error(f"WhatsApp डिलीवरी त्रुटि: {e}")
        print(f"\n[WARNING] WhatsApp संदेश नहीं भेजा जा सका: {e}")
    finally:
        if wa_driver:
            try:
                wa_driver.quit()
            except Exception:
                pass


# =========================================================
# 17. Main Runner Loop
# =========================================================
def main():
    print("\n[START] स्वामित्व डैशबोर्ड एवं डेटाबेस ऑटोमेशन शुरू किया जा रहा है...\n")
    logging.info("Svamitva Dashboard All-in-One Automation shuru...")

    init_database()

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": BASE_DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })

    print("---> Chrome WebDriver शुरू हो रहा है...")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": BASE_DOWNLOAD_DIR
    })

    df_state_comp = None
    df_state_pct = None

    try:
        # Step 1: State & Gwalior Division Reports (Only 55 Districts)
        print("\n---> चरण 1: राज्य स्तरीय 55 जिलों का डेटा प्राप्त किया जा रहा है...")
        try:
            clean_df, district_col, comp_col, pct_col = get_state_districts_data(driver)

            if comp_col:
                r1_excel = os.path.join(OUTPUT_DIR, "State_Districts_Ranked_By_Completion.xlsx")
                r1_pdf = os.path.join(OUTPUT_DIR, "State_Districts_Ranked_By_Completion.pdf")
                generate_district_ranked_report(driver, clean_df, district_col, comp_col, "कम्प्लीशन", r1_excel, r1_pdf)
                df_state_comp = clean_df.sort_values(by=comp_col, ascending=False).reset_index(drop=True)
                print("  ✓ State Completion रिपोर्ट (55 जिले + Total Row) तैयार।")

            if pct_col:
                r2_excel = os.path.join(OUTPUT_DIR, "State_Districts_Ranked_By_Percentage.xlsx")
                r2_pdf = os.path.join(OUTPUT_DIR, "State_Districts_Ranked_By_Percentage.pdf")
                generate_district_ranked_report(driver, clean_df, district_col, pct_col, "प्रतिशत", r2_excel, r2_pdf)
                df_state_pct = clean_df.sort_values(by=pct_col, ascending=False).reset_index(drop=True)
                print("  ✓ State Percentage रिपोर्ट (55 जिले + Total Row) तैयार।")

            if pct_col:
                pattern = "|".join(GWALIOR_DIVISION_DISTRICTS)
                gwalior_df = clean_df[
                    clean_df[district_col].astype(str).str.contains(pattern, case=False, na=False)].copy()
                gw_excel = os.path.join(OUTPUT_DIR, "Gwalior_Division_Ranked_By_Percentage.xlsx")
                gw_pdf = os.path.join(OUTPUT_DIR, "Gwalior_Division_Ranked_By_Percentage.pdf")
                generate_district_ranked_report(driver, gwalior_df, district_col, pct_col, "प्रतिशत", gw_excel, gw_pdf,
                                                custom_division_title="Gwalior Division")
                print("  ✓ Gwalior Division रिपोर्ट तैयार।")
        except Exception as e:
            print(f"[ERROR] राज्य स्तरीय रैंकिंग में त्रुटि: {e}")
            traceback.print_exc()

        # Step 2: District Ashoknagar Reports
        print("\n---> चरण 2: जिला अशोकनगर की तहसील एवं पटवारीवार रिपोर्ट्स बनाई जा रही हैं...")

        district_goshwara_rows = []
        district_zero_patwaris_list = []
        all_tehsil_patwaris_list = []
        low_progress_summary_list = []

        global_pct_col = "प्रतिशत"
        global_tot_col = None
        global_comp_col = None

        for s_no, (tehsil_name, url) in enumerate(TEHSILS_DATA.items(), start=1):
            print(f"  [{s_no}/8] प्रोसेस की जा रही है: तहसील {tehsil_name}...")

            try:
                driver.get(url)
            except Exception as e:
                print(f"    ! URL नहीं खुला: {tehsil_name} -> {e}")
                continue

            raw_df = get_tehsil_dataframe(driver, tehsil_name)
            if raw_df is None or raw_df.empty:
                print(f"    ! डेटा नहीं मिला: {tehsil_name}")
                continue

            clean_patwaris_df, patwari_summary_df, goshwara_row, zero_patwaris, pct_col_name, tot_col, comp_col, count_1_to_9, count_zero = process_tehsil_data(
                raw_df, tehsil_name)

            global_pct_col = pct_col_name
            global_tot_col = tot_col
            global_comp_col = comp_col

            goshwara_row["क्र."] = s_no
            district_goshwara_rows.append(goshwara_row)
            all_tehsil_patwaris_list.append(clean_patwaris_df)

            low_progress_summary_list.append({
                "तहसील का नाम": tehsil_name,
                "10 से कम प्रगति": count_1_to_9,
                "0 प्रगति": count_zero
            })

            if not zero_patwaris.empty:
                district_zero_patwaris_list.append(zero_patwaris)

            t_excel = os.path.join(OUTPUT_DIR, f"Svamitva_{tehsil_name}_PatwariSummary.xlsx")
            saved_excel_path = safe_save_excel(patwari_summary_df, t_excel)
            format_excel_file(saved_excel_path, pct_col_name)

            t_pdf = os.path.join(OUTPUT_DIR, f"Svamitva_{tehsil_name}_PatwariSummary.pdf")
            generate_pdf_via_chrome(driver, patwari_summary_df, t_pdf,
                                    f"स्वामित्व योजना पटवारीवार प्रगति (अवरोही क्रम) - तहसील {tehsil_name}",
                                    pct_col_name)

        # 10 Se Kam & 0 Progress Report
        if low_progress_summary_list:
            low_prog_excel = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Less_Than_10_and_Zero_Progress.xlsx")
            low_prog_pdf = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Less_Than_10_and_Zero_Progress.pdf")
            generate_low_progress_report(driver, low_progress_summary_list, low_prog_excel, low_prog_pdf)
            print("  ✓ 10 से कम एवं 0 प्रगति सारांश तैयार।")

        # All Patwaris Report (<50 Highlight + Separate Pages)
        if all_tehsil_patwaris_list and global_tot_col and global_comp_col:
            print("\n---> समस्त पटवारी तहसीलवार (<50 अलर्ट एवं सेपरेट पेज) रिपोर्ट बनाई जा रही है...")
            all_pat_excel = os.path.join(OUTPUT_DIR, "District_Ashoknagar_All_Patwaris_Tehsilwise.xlsx")
            all_pat_pdf = os.path.join(OUTPUT_DIR, "District_Ashoknagar_All_Patwaris_Tehsilwise.pdf")
            generate_all_patwaris_tehsilwise_custom_report(
                driver, all_tehsil_patwaris_list, global_tot_col, global_comp_col, global_pct_col, all_pat_excel,
                all_pat_pdf
            )
            print("  ✓ District_Ashoknagar_All_Patwaris_Tehsilwise रिपोर्ट तैयार।")

        # Tehsilwise Bottom 10 Patwaris Report
        if all_tehsil_patwaris_list and global_tot_col and global_comp_col:
            print("---> प्रत्येक तहसील के न्यूनतम 10 पटवारी रिपोर्ट बनाई जा रही है...")
            bottom_10_excel = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Tehsilwise_Bottom_10_Patwaris.xlsx")
            bottom_10_pdf = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Tehsilwise_Bottom_10_Patwaris.pdf")
            generate_tehsilwise_bottom_10_patwaris_report(driver, all_tehsil_patwaris_list, global_tot_col,
                                                          global_comp_col, bottom_10_excel, bottom_10_pdf)
            print("  ✓ District_Ashoknagar_Tehsilwise_Bottom_10_Patwaris रिपोर्ट तैयार।")

        # District Tehsil Goshwara & Comparison Reports
        if district_goshwara_rows:
            print("---> जिला गोशवारा एवं तुलनात्मक रिपोर्ट्स बनाई जा रही हैं...")
            goshwara_df = pd.DataFrame(district_goshwara_rows)
            if global_pct_col in goshwara_df.columns:
                goshwara_df = goshwara_df.sort_values(by=global_pct_col, ascending=False).reset_index(drop=True)
                goshwara_df["क्र."] = range(1, len(goshwara_df) + 1)

            top_pct_row = goshwara_df.loc[goshwara_df[global_pct_col].idxmax()]
            top_comp_row = goshwara_df.loc[goshwara_df[
                global_comp_col].idxmax()] if global_comp_col and global_comp_col in goshwara_df.columns else None

            top_box_html = f"🏆 <b>सर्वश्रेष्ठ प्रदर्शनकारी तहसील (Top Performing Tehsils):</b><br>"
            top_box_html += f"• <b>प्रतिशत में प्रथम:</b> तहसील <u>{top_pct_row['तहसील']}</u> ({top_pct_row[global_pct_col]}%)<br>"
            if top_comp_row is not None:
                top_box_html += f"• <b>कम्प्लीशन (पूर्ण संख्या) में प्रथम:</b> तहसील <u>{top_comp_row['तहसील']}</u> ({int(top_comp_row[global_comp_col])} पूर्ण Plots)"

            dist_total_row = {c: "" for c in goshwara_df.columns}
            dist_total_row["क्र."] = ""
            dist_total_row["तहसील"] = "संपूर्ण जिला महायोग (Grand Total)"
            dist_total_row["कुल पटवारी"] = goshwara_df["कुल पटवारी"].sum()
            dist_total_row["कुल ग्राम"] = goshwara_df["कुल ग्राम"].sum()

            for c in goshwara_df.select_dtypes(include=['number']).columns:
                if c not in ["क्र.", "कुल पटवारी", "कुल ग्राम", global_pct_col]:
                    dist_total_row[c] = goshwara_df[c].sum()

            if global_comp_col and global_tot_col and dist_total_row[global_tot_col] > 0:
                dist_total_row[global_pct_col] = round(
                    (dist_total_row[global_comp_col] / dist_total_row[global_tot_col] * 100), 2)
            else:
                dist_total_row[global_pct_col] = 0.0

            final_goshwara_df = pd.concat([goshwara_df, pd.DataFrame([dist_total_row])], ignore_index=True)

            dist_excel = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Tehsil_Goshwara.xlsx")
            saved_dist_excel = safe_save_excel(final_goshwara_df, dist_excel)
            format_excel_file(saved_dist_excel, global_pct_col)

            init_col = find_column_by_keywords(final_goshwara_df.columns,
                                               ["initiated", "प्रारंभ", "रजिस्ट्रीप्रारंभ", "पंजीयनप्रारंभ"])
            veri_col = find_column_by_keywords(final_goshwara_df.columns, ["verification", "सत्यापन", "पटवारीसत्यापन"])
            subm_col = find_column_by_keywords(final_goshwara_df.columns, ["submitted", "सबमिट", "जमा"])

            init_val = dist_total_row.get(init_col, "0") if init_col else "0"
            veri_val = dist_total_row.get(veri_col, "0") if veri_col else "0"
            subm_val = dist_total_row.get(subm_col, "0") if subm_col else "0"

            goshwara_footer_html = f"""
            <b>1. Registry Initiated :</b> <span style="color: #0d47a1; font-weight: bold;">{init_val}</span> (Pending at Patwari Level Initiated by SAARA App through Patwari or By MP E-Seva App)<br>
            <b>2. Patwari Verification :</b> <span style="color: #0d47a1; font-weight: bold;">{veri_val}</span> (Pending At patwari Level for Verification)<br>
            <b>3. Submitted :</b> <span style="color: #0d47a1; font-weight: bold;">{subm_val}</span> (Pending at Tehsildar level for Completion From Sampda App)
            """

            dist_pdf = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Tehsil_Goshwara.pdf")
            generate_pdf_via_chrome(driver, final_goshwara_df, dist_pdf,
                                    "स्वामित्व योजना प्रगति गोशवारा (प्रतिशत अनुसार अवरोही क्रम) - जिला अशोकनगर",
                                    global_pct_col, footer_notes_html=goshwara_footer_html,
                                    top_performer_box=top_box_html)
            print("  ✓ District_Ashoknagar_Tehsil_Goshwara तैयार।")

            # -----------------------------------------------------
            # तुलनात्मक रिपोर्ट्स जनरेशन (Morning Baseline, 2-Hour Ago & Completed)
            # -----------------------------------------------------
            print("\n---> तुलनात्मक रिपोर्ट्स तैयार की जा रही हैं...")

            # 1. केवल Completed प्लॉट्स तुलनात्मक रिपोर्ट
            comp_only_excel = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Tehsilwise_Completed_Comparison.xlsx")
            comp_only_pdf = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Tehsilwise_Completed_Comparison.pdf")
            generate_tehsilwise_completed_comparison_report(driver, goshwara_df, global_tot_col, global_comp_col,
                                                            comp_only_excel, comp_only_pdf)
            print("  ✓ 1. District_Ashoknagar_Tehsilwise_Completed_Comparison तैयार।")

            # 2. Daily Morning Baseline Comparison
            compare_excel = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Comparative_Progress.xlsx")
            compare_pdf = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Comparative_Progress.pdf")
            generate_comparative_growth_report(driver, goshwara_df, global_tot_col, global_comp_col, global_pct_col,
                                               compare_excel, compare_pdf)
            print("  ✓ 2. District_Ashoknagar_Comparative_Progress (Morning Baseline) तैयार।")

            # 3. NEW: 2-Hour Ago Comparison Report
            two_hour_excel = os.path.join(OUTPUT_DIR, "District_Ashoknagar_2Hour_Comparison.xlsx")
            two_hour_pdf = os.path.join(OUTPUT_DIR, "District_Ashoknagar_2Hour_Comparison.pdf")
            generate_2hour_ago_comparison_report(driver, goshwara_df, global_tot_col, global_comp_col, global_pct_col,
                                                 two_hour_excel, two_hour_pdf)
            print("  ✓ 3. District_Ashoknagar_2Hour_Comparison (2 घंटे पूर्व से तुलना) तैयार।")

            # अब आज का नया डेटा डेटाबेस में सुरक्षित करें
            print("---> डेटाबेस में आज का नया डेटा सिंक किया जा रहा है...")
            save_today_data_to_db(goshwara_df, all_tehsil_patwaris_list, global_tot_col, global_comp_col,
                                  global_pct_col)

            # Executive One-Pager Summary Generation (With Names of Patwaris)
            if all_tehsil_patwaris_list:
                all_p_concat = pd.concat(all_tehsil_patwaris_list, ignore_index=True)
                one_pager_excel = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Executive_OnePager_Summary.xlsx")
                one_pager_pdf = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Executive_OnePager_Summary.pdf")
                generate_one_pager_executive_summary(driver, goshwara_df, all_p_concat, dist_total_row, global_pct_col,
                                                     global_comp_col, one_pager_excel, one_pager_pdf)
                print("  ✓ Executive One-Pager Summary (पटवारियों के नाम सहित) तैयार।")

            # PowerPoint Presentation (.pptx)
            print("---> PowerPoint Presentation (.pptx) तैयार की जा रही है...")
            review_pptx_path = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Svamitva_Review.pptx")
            generate_review_pptx(
                review_pptx_path, df_state_comp, df_state_pct, goshwara_df,
                all_tehsil_patwaris_list, global_comp_col, global_tot_col, global_pct_col
            )
            print("  ✓ PowerPoint Presentation (.pptx) तैयार।")

        # Zero Progress Patwaris
        if district_zero_patwaris_list:
            all_zero_df = pd.concat(district_zero_patwaris_list, ignore_index=True)
            all_zero_df.insert(0, "क्र.", range(1, len(all_zero_df) + 1))
            zero_excel = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Zero_Progress_Patwaris.xlsx")
            saved_zero_excel = safe_save_excel(all_zero_df, zero_excel)
            format_excel_file(saved_zero_excel, global_pct_col, is_zero_report=True)

            zero_pdf = os.path.join(OUTPUT_DIR, "District_Ashoknagar_Zero_Progress_Patwaris.pdf")
            generate_pdf_via_chrome(driver, all_zero_df, zero_pdf,
                                    "शून्य प्रगति (0 Complete) वाले पटवारियों की सूची - जिला अशोकनगर", global_pct_col,
                                    is_zero_report=True)
            print("  ✓ शून्य प्रगति रिपोर्ट तैयार।")

    except Exception as e:
        print(f"\n[CRITICAL ERROR] ऑटोमेशन के दौरान त्रुटि आई: {e}")
        traceback.print_exc()
    finally:
        driver.quit()
        print(f"\n=======================================================")
        print(f"सभी फ़ाइलें सफलतापूर्वक यहाँ सुरक्षित की गईं:\n{OUTPUT_DIR}")
        print(f"डेटाबेस फ़ाइल: {DB_PATH}")
        print(f"=======================================================\n")

        # =====================================================
        # Step 4: WhatsApp Auto Delivery
        # =====================================================
        TARGET_WHATSAPP_NUMBER = "91XXXXXXXXXX"  # अपना 10 अंकों का WhatsApp नंबर यहाँ लिखें (जैसे: "919876543210")

        reports_to_send = [
            os.path.join(OUTPUT_DIR, "District_Ashoknagar_Comparative_Progress.pdf"),
            os.path.join(OUTPUT_DIR, "District_Ashoknagar_2Hour_Comparison.pdf"),
            os.path.join(OUTPUT_DIR, "District_Ashoknagar_Tehsilwise_Completed_Comparison.pdf"),
            os.path.join(OUTPUT_DIR, "District_Ashoknagar_Executive_OnePager_Summary.pdf"),
            os.path.join(OUTPUT_DIR, "District_Ashoknagar_Tehsilwise_Bottom_10_Patwaris.pdf"),
            os.path.join(OUTPUT_DIR, "District_Ashoknagar_All_Patwaris_Tehsilwise.pdf"),
            os.path.join(OUTPUT_DIR, "District_Ashoknagar_Tehsil_Goshwara.pdf")
        ]

        summary_msg = (
            f"📢 *स्वामित्व योजना दैनिक समीक्षा एवं तुलनात्मक बुलेटिन - जिला अशोकनगर*\n"
            f"📅 दिनांक व समय: {get_current_time_str()}\n\n"
            f"✅ *मॉर्निंग बेसलाइन तुलनात्मक वृद्धि* रिपोर्ट संलग्न है।\n"
            f"✅ *2 घंटे पूर्व से तुलनात्मक प्रगति* रिपोर्ट संलग्न है।\n"
            f"✅ Executive One-Pager Summary (पटवारियों के नाम सहित) संलग्न है।\n"
            f"✅ प्रत्येक तहसील के न्यूनतम 10 पटवारी (Bottom 10) रिपोर्ट संलग्न है।"
        )

        try:
            if TARGET_WHATSAPP_NUMBER and "X" not in TARGET_WHATSAPP_NUMBER:
                print("---> WhatsApp पर संदेश भेजा जा रहा है...")
                send_whatsapp_report(TARGET_WHATSAPP_NUMBER, summary_msg, reports_to_send, is_group=False)
            else:
                print("---> WhatsApp नंबर सेट नहीं है, संदेश छोड़ दिया गया।")
        except Exception as e:
            print(f"[ERROR] WhatsApp भेजने में समस्या: {e}")

        try:
            os.startfile(OUTPUT_DIR)
        except Exception:
            pass


if __name__ == "__main__":
    main()