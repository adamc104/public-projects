"""One reader per state. Each returns {"YYYY-MM": enrollment} for every month it can find.

Each reader is independent: if a state changes its website, only that reader breaks.
The row each reader pulls was matched against reference historicals (see reference.csv).
"""
import json
import re
from urllib.parse import urljoin

import pandas as pd

from helpers import (Browser, DOWNLOADS, add_months, cached_download, clean, http_get, month_key,
                     month_number, numbers_in, pdf_lines, to_int)

START = "2023-01"  # backfill from here


# ---------------------------------------------------------------- Florida
# Monthly Excel. The latest file's "MEDICAID" sheet carries every month back to 2015.

def florida(browser):
    base = "https://ahca.myflorida.com"
    page_url = (f"{base}/medicaid/medicaid-finance-and-analytics/medicaid-data-analytics/"
                "medicaid-monthly-enrollment-report.html")
    page = http_get(page_url).text
    latest = sorted(set(re.findall(r"/file/medicaid/ENR_(\d{6})\.xls", page)))[-1]
    out = _florida_file(cached_download("FL", f"ENR_{latest}.xls", f"{base}/file/medicaid/ENR_{latest}.xls"))

    # The latest file skips some older months (e.g. 2023); fill those from each month's own file
    older = {}
    for href in re.findall(r'href="([^"]*ENR_[^"]*\.xls)"', page):
        m = re.search(r"ENR_(\d{4})-?(\d{2})\.xls", href)
        if m:
            older[month_key(m.group(1), m.group(2))] = urljoin(page_url, href)
    key = START
    while key < max(out):
        if key not in out and key in older:
            out.update({k: v for k, v in _florida_file(
                cached_download("FL", f"ENR_{key}.xls", older[key])).items() if k not in out})
        key = add_months(key, 1)
    return {k: v for k, v in out.items() if k >= START}


def _florida_file(path):
    sheet = pd.read_excel(path, sheet_name="MEDICAID", header=None)
    header_row = sheet.index[sheet[0].astype(str).str.strip() == "PLAN_NAME"][0]
    total_row = sheet.index[sheet[0].astype(str).str.strip() == "FLORIDA_MEDICAID ENROLLMENT TOTAL"][0]
    out = {}
    for col in range(1, sheet.shape[1]):
        label = str(sheet.iat[header_row, col])  # e.g. "AUGUST_2026"
        m = re.fullmatch(r"([A-Z]+)_(\d{4})", label.strip())
        value = to_int(sheet.iat[total_row, col])
        if m and value:
            out[month_key(m.group(2), month_number(m.group(1)))] = value
    return out


# ---------------------------------------------------------------- Kansas
# One PDF per state fiscal year (July-June), "Medical Assistance Report". Needs the browser.

def kansas(browser):
    url = "https://www.kancare.ks.gov/policies-and-reports/medical-assistance-report"
    links = [(t, h) for t, h in browser.links(url) if re.match(r"MAR FY \d{4}", t)]
    first_fy = int(START[:4])  # FY2023 = Jul 2022 - Jun 2023
    out = {}
    for text, href in links:
        fy = int(re.match(r"MAR FY (\d{4})", text).group(1))
        if fy < first_fy:
            continue
        is_current = links[0][1] == href  # the newest file is updated in place each month
        path = cached_download("KS", f"MAR_FY{fy}.pdf", fetch=lambda: browser.fetch_bytes(href),
                               refresh=is_current)
        for line in pdf_lines(path):
            if line.startswith("Total "):  # first "Total" line = all beneficiaries, Jul..Jun + average
                values = numbers_in(line)[:12]
                for i, v in enumerate(values):
                    if v:
                        out[add_months(month_key(fy - 1, 7), i)] = v
                break
    return {k: v for k, v in out.items() if k >= START}


# ---------------------------------------------------------------- Kentucky
# One PDF per month ("Monthly Membership Counts by County"), listed behind a year dropdown on the main
# page (current year) and an archive page (2024-2025; 2023 isn't online). Archive file names are
# inconsistent, so the month comes from the PDF's own "Coverage Month" line. Total is the "Sum:" line.

def kentucky(browser):
    pages = [
        "https://chfs.ky.gov/agencies/dms/dafm/Pages/statistics.aspx",
        "https://chfs.ky.gov/agencies/dms/dafm/Pages/statistics-archive.aspx",
    ]
    hrefs = set()
    for url in pages:
        browser.open(url, wait_ms=4000)
        dropdown = browser.page.locator("select[name$='ddlViewSelect']").first
        years = [o.strip() for o in dropdown.locator("option").all_inner_texts() if re.match(r"\d{4}", o.strip())]
        for year in years:
            if int(year[:4]) < int(START[:4]):
                continue
            browser.page.locator("select[name$='ddlViewSelect']").first.select_option(label=year)
            browser.page.wait_for_timeout(5000)
            hrefs |= {h for _, h in browser.links() if "/dms/stats/" in h and h.lower().endswith(".pdf")}
    out = {}
    for href in sorted(hrefs):
        path = cached_download("KY", href.rsplit("/", 1)[-1], href)
        lines = pdf_lines(path, pages=[0])
        coverage = next((l for l in lines if re.match(r"\d{1,2}/1/\d{4} \d{1,2}/\d{1,2}/\d{4}", l)), None)
        total = next((l for l in lines if l.startswith("Sum:")), None)
        if coverage and total:
            m, _, y = coverage.split()[0].split("/")
            out[month_key(y, m)] = numbers_in(total)[-1]
    return {k: v for k, v in out.items() if k >= START}


# ---------------------------------------------------------------- Minnesota
# One Excel per month (managed care enrollment). Total is the "Total Enrollment" row. Needs the browser.

def minnesota(browser):
    url = ("https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION"
           "&RevisionSelectionMethod=LatestReleased&dDocName=dhs16_141529")
    out = {}
    for text, href in browser.links(url, wait_ms=5000):
        m = re.search(r"enrollment figures, (\w+) (\d{4}) \(XLSX\)", text, re.I)
        if not m:
            continue
        key = month_key(m.group(2), month_number(m.group(1)))
        if key < START:
            continue
        path = cached_download("MN", f"MN_{key}.xlsx", fetch=lambda: browser.fetch_bytes(href))
        for name, sheet in pd.read_excel(path, sheet_name=None, header=None).items():
            rows = sheet.index[sheet[0].astype(str).str.strip().str.lower() == "total enrollment"]
            if len(rows):
                out[key] = max(v for v in map(to_int, sheet.iloc[rows[0]].tolist()) if v)
                break
    return out


# ---------------------------------------------------------------- Missouri
# One running history Excel ("Caseload Counter"), back to 2004. Row "MO HealthNet Enrollees".

def missouri(browser):
    page = http_get("https://dss.mo.gov/media/file/caseload-counter-historical-information").text
    href = re.search(r'href="([^"]+\.xlsx)"', page).group(1)
    path = cached_download("MO", "caseload_counter_latest.xlsx", href, refresh=True)
    sheet = pd.read_excel(path, header=None)
    header_row = 3
    row = sheet.index[sheet[0].astype(str).str.startswith("MO HealthNet Enrollees")][0]
    out = {}
    for col in range(1, sheet.shape[1]):
        when, value = sheet.iat[header_row, col], to_int(sheet.iat[row, col])
        if hasattr(when, "year") and value:
            out[month_key(when.year, when.month)] = value
    return {k: v for k, v in out.items() if k >= START}


# ---------------------------------------------------------------- New Hampshire
# One PDF, replaced monthly. Page 2: last 13 months of "Grand Total Full Medicaid". Needs the browser.

def new_hampshire(browser):
    import pdfplumber
    links = browser.links("https://www.dhhs.nh.gov/ombp/medicaid/enrollment-data.htm")
    href = next(h for t, h in links if t == "Medicaid Enrollment Data")
    path = cached_download("NH", "nh_medicaid_enrollment.pdf", fetch=lambda: browser.fetch_bytes(href),
                           refresh=True)
    out, history = {}, []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            is_history = (page.extract_text() or "").startswith("Historical")
            for table in page.extract_tables():
                for row in table:
                    if not (row and row[0] and clean(row[0]) == "Grand Total Full Medicaid"):
                        continue
                    if is_history:  # one row per state fiscal year (Jul-Jun), oldest first
                        history.append([v for v in map(to_int, row[1:]) if v])
                        continue
                    for label, value in zip(table[0][1:], row[1:]):
                        # Current-period table stores dates backwards, e.g. "6202/13/8" is 8/31/2026
                        m = re.fullmatch(r"(\d{1,2})/\d{1,2}/(\d{4})", clean(label or "")[::-1])
                        if m and to_int(value):
                            out[month_key(m.group(2), m.group(1))] = to_int(value)

    # History rows aren't labelled with dates in a readable way, so line the newest row up against the
    # current-period months: try each fiscal year it could be and keep the one where the numbers agree.
    latest_year = int(max(out)[:4])
    for last_fy in (latest_year, latest_year - 1, latest_year + 1):
        mapped = {}
        for i, values in enumerate(reversed(history)):
            fy = last_fy - i
            for j, v in enumerate(values):
                mapped[add_months(month_key(fy - 1, 7), j)] = v
        overlap = [k for k in out if k in mapped]
        if overlap and all(mapped[k] == out[k] for k in overlap):
            out = {**mapped, **out}
            break
    return {k: v for k, v in out.items() if k >= START}


# ---------------------------------------------------------------- New York
# One Excel per month at a predictable address. "Medicaid Managed Care Table" -> first "State Total" row,
# last column (same number as the NYSOH table's state total, which some 2023 files leave out).

def new_york(browser, through):
    base = "https://www.health.ny.gov/health_care/managed_care/reports/enrollment/monthly"
    out = {}
    key = START
    while key <= through:
        y, m = key.split("-")
        name = f"en{m}_{y[2:]}.xlsx"
        try:
            path = cached_download("NY", name, f"{base}/{y}/docs/{name}")
        except Exception:  # month not published yet
            key = add_months(key, 1)
            continue
        for tab in ("Medicaid Managed Care Table", "NYSOH Table"):  # some months only have it on one tab
            sheet = pd.read_excel(path, sheet_name=tab, header=None)
            rows = sheet.index[sheet[0].astype(str).str.strip().str.upper() == "STATE TOTAL"]
            if len(rows):
                out[key] = [v for v in map(to_int, sheet.iloc[rows[0]].tolist()) if v][-1]
                break
        key = add_months(key, 1)
    return out


# ---------------------------------------------------------------- North Carolina
# History: "Monthly Medicaid Enrollees" Excel (column "MEDICAID + MED. EXPANSION"), published a few times a year.
# Newer months: the Tableau enrollment dashboard, read by OCR with a check that the plan tiles sum to the total.

NC_VIEW = ("https://bidashboards.nc.gov/t/DHHS-Medicaid/views/NewMCEnroll-Website-Revised-v4/"
           "NCMedicaidEnrollmentOverview?:embed=y&:showVizHome=no")


def north_carolina(browser, through):
    page = http_get("https://medicaid.ncdhhs.gov/reports/nc-medicaid-enrollment-reports").text
    href = re.search(r'href="([^"]*monthly-medicaid-enrollees-sfy\d{4}[^"]*)"', page, re.I).group(1)
    if href.startswith("/"):
        href = "https://medicaid.ncdhhs.gov" + href
    path = cached_download("NC", "nc_monthly_enrollees.xlsx", href, refresh=True)
    out = {}
    for name, sheet in pd.read_excel(path, sheet_name=None, header=None).items():
        if not name.upper().startswith("SFY"):
            continue
        # Total column: "MEDICAID + MED. EXPANSION" since Dec 2023, "MEDICAID + COVID-19" before. Take the rightmost.
        cols = [c for c in range(sheet.shape[1]) if "MEDICAID +" in " ".join(map(str, sheet.iloc[:8, c]))]
        if not cols:
            continue
        col = cols[-1]
        for r in range(sheet.shape[0]):
            year, mon, value = sheet.iat[r, 0], sheet.iat[r, 1], to_int(sheet.iat[r, col])
            if to_int(year) and isinstance(mon, str) and value:
                out[month_key(to_int(year), month_number(mon))] = value

    # Months after the Excel file: read them off the dashboard (cached so each month is only read once)
    cache_file = DOWNLOADS / "NC" / "dashboard_reads.json"
    cache = json.loads(cache_file.read_text()) if cache_file.exists() else {}
    key = add_months(max(out), 1)
    while key <= through:
        if key not in cache:
            value = _nc_dashboard_month(browser, key)
            if value is None:
                break  # dashboard doesn't have this month yet
            cache[key] = value
            cache_file.parent.mkdir(parents=True, exist_ok=True)
            cache_file.write_text(json.dumps(cache, indent=1))
        key = add_months(key, 1)
    out.update({k: v for k, v in cache.items() if k not in out})
    return {k: v for k, v in out.items() if k >= START}


def _nc_dashboard_month(browser, key):
    import pytesseract
    from PIL import Image, ImageOps
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    from helpers import MONTHS
    y, m = key.split("-")
    label = f"{MONTHS[int(m) - 1].capitalize()} {y}"
    page = browser.page
    if "bidashboards" not in page.url:
        browser.open(NC_VIEW, wait_ms=0)
        page.get_by_text("Total").first.wait_for(timeout=90000)
        page.wait_for_timeout(5000)
    page.locator(".tabComboBoxNameContainer, .tabComboBox").first.click()
    option = page.get_by_text(label, exact=True)
    if not option.count():
        page.keyboard.press("Escape")
        return None
    option.last.click()
    page.wait_for_timeout(10000)
    if page.locator(".tabComboBoxName").first.inner_text().strip() != label:
        raise RuntimeError(f"NC dashboard did not switch to {label}")
    shot = DOWNLOADS / "NC" / f"dashboard_{key}.png"
    shot.parent.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(shot))
    band = Image.open(shot).crop((90, 300, 1180, 340)).convert("L")  # the row of six number tiles
    band = ImageOps.autocontrast(band.resize((band.width * 3, band.height * 3)))
    text = pytesseract.image_to_string(band, config="--psm 6 -c tessedit_char_whitelist=0123456789,")
    nums = [to_int(n) for n in re.findall(r"\d[\d,]{2,}", text)]
    # Safety check: the five plan tiles must add up exactly to the Total Enrollment tile
    if len(nums) != 6 or nums[0] != sum(nums[1:]):
        raise RuntimeError(f"NC OCR check failed for {label}: read {nums} (see {shot.name})")
    return nums[0]


# ---------------------------------------------------------------- West Virginia
# One PDF per calendar year. Page 1 table: the last "Total" row is total Medicaid.

def west_virginia(browser):
    import pdfplumber
    from bs4 import BeautifulSoup
    page = BeautifulSoup(http_get("https://bms.wv.gov/managed-care-reports/mco-monthly-enrollment").text,
                         "html.parser")
    years = {}
    for a in page.find_all("a", href=True):
        text = clean(a.get_text())
        if re.fullmatch(r"\d{4}", text) and "/media/" in a["href"]:
            years.setdefault(int(text), "https://bms.wv.gov" + a["href"] if a["href"].startswith("/") else a["href"])
    out = {}
    newest = max(years)
    for year, href in years.items():
        if year < int(START[:4]):
            continue
        path = cached_download("WV", f"WV_{year}.pdf", href, refresh=(year == newest))
        with pdfplumber.open(path) as pdf:
            table = pdf.pages[0].extract_tables()[0]
        header = [clean(c or "") for c in table[0]]
        total = [r for r in table if r and clean(r[0] or "") == "Total"][-1]
        for label, value in zip(header[1:], total[1:]):
            if label and to_int(value):
                out[month_key(year, month_number(label))] = to_int(value)
    return out


# ---------------------------------------------------------------- Wisconsin
# One PDF per month at a predictable address, each with the last six months.
# "Total" row; the last number is the file's own month.

def wisconsin(browser, through):
    from helpers import MONTHS
    out = {}
    key = START
    while key <= through:  # oldest first, so newer files overwrite any revised months
        y, m = key.split("-")
        name = f"bcpstate-{MONTHS[int(m) - 1]}{y}.pdf"
        try:
            path = cached_download("WI", name, f"https://www.dhs.wisconsin.gov/badgercareplus/{name}")
        except Exception:  # month not published yet
            key = add_months(key, 1)
            continue
        # A header line names the six months ("Feb 2023 ... Jul 2023" in older files, "MAR2026 ... AUG2026" in newer);
        # the next "Total" line has the numbers
        lines = pdf_lines(path)
        if sum(l.startswith("Total ") for l in lines) > 3:
            # Mis-posted county report (e.g. "november2023" is by county); later files cover these months
            key = add_months(key, 1)
            continue
        for i, line in enumerate(lines):
            months = re.findall(r"\b([A-Za-z]{3}) ?(\d{4})\b", line)
            if len(months) == 6:
                total = next(l for l in lines[i + 1:] if l.startswith("Total "))
                for (mon, year), v in zip(months, numbers_in(total)):
                    out[month_key(year, month_number(mon))] = v
                break
        key = add_months(key, 1)
    return {k: v for k, v in out.items() if k >= START}


# ---------------------------------------------------------------- Mississippi
# Site under maintenance (Oct 2026). Use manual_entries.csv until it's back.

def mississippi(browser):
    return {}  # no reader yet - the tracker falls back to manual_entries.csv


STATES = {
    "Florida": florida,
    "Kansas": kansas,
    "Kentucky": kentucky,
    "Minnesota": minnesota,
    "Mississippi": mississippi,
    "Missouri": missouri,
    "New Hampshire": new_hampshire,
    "New York": new_york,
    "North Carolina": north_carolina,
    "West Virginia": west_virginia,
    "Wisconsin": wisconsin,
}
