import streamlit as st
import asyncio
import time
import os
import csv
import io
import zipfile
from pathlib import Path
from datetime import datetime
import subprocess
import threading
import win32com.client
import win32process
import pythoncom
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from dotenv import load_dotenv

load_dotenv()


# =========================================================
# DOWNLOAD DIRECTORY - MPHASIS
# =========================================================
DOWNLOAD_DIR = Path(r"C:\Gen AI\Mphasis\Downloads")
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

# Final output file name (was "cleaned_tracker.xlsx"). get_latest_excel_file()
# excludes this exact name (by stem) from candidate source files — keep that
# exclusion in sync with this if it's ever renamed again.
CLEANED_FILE_NAME = "Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx"

TRIGGER_URL = "https://arsmistracker.authbridge.app/tracker/generate_excel_data/5973?email=true"

# =========================================================
# MIS EXPORT QUERY TOOL - LOGIN & QUERY CONFIG
# =========================================================
MIS_LOGIN_URL = os.getenv("MIS_LOGIN_URL", "https://mis.authbridge.com/export_query/login.php")
MIS_USERNAME = os.getenv("MIS_USERNAME")
MIS_PASSWORD = os.getenv("MIS_PASSWORD")
MIS_HOST = "Bridge Live"
MIS_DATABASE = "Bridge Live"

MIS_EXPORT_DIR = DOWNLOAD_DIR / "MIS_Exports"
MIS_EXPORT_DIR.mkdir(parents=True, exist_ok=True)


# =========================================================
# STREAMLIT CONFIG
# =========================================================
st.set_page_config(page_title="Mphasis Tracker", layout="centered", page_icon="📧")
st.title("📧 Mphasis Tracker")


# =========================================================
# THREAD-SAFE COM CONTEXT
# =========================================================
class COMContext:
    """Thread-safe COM initialization"""
    _lock = threading.Lock()

    @staticmethod
    def initialize():
        with COMContext._lock:
            try:
                pythoncom.CoInitialize()
                return True
            except:
                return False

    @staticmethod
    def uninitialize():
        with COMContext._lock:
            try:
                pythoncom.CoUninitialize()
            except:
                pass


# =========================================================
# CHECK OUTLOOK PROCESS
# =========================================================
def is_outlook_running():
    """Check if Outlook process is running"""
    try:
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq outlook.exe'],
                              capture_output=True, text=True, timeout=5)
        return 'outlook.exe' in result.stdout.lower()
    except:
        return False


# =========================================================
# FUNCTION TO SEARCH AND DOWNLOAD LATEST EMAIL ATTACHMENT
# =========================================================
def get_target_inbox(namespace, account_name=""):
    """Return the Inbox to search. With multiple Outlook accounts configured,
    GetDefaultFolder(6) only returns the DEFAULT account's Inbox, which may not
    be the one the tracker email actually lands in. If account_name is given,
    find the matching top-level account folder and use its Inbox instead."""
    if not account_name.strip():
        return namespace.GetDefaultFolder(6)

    for folder in namespace.Folders:
        if account_name.strip().lower() in folder.Name.lower():
            return folder.Folders["Inbox"]

    return None


def download_latest_from_email(subject_keyword, wait_minutes=10, account_name=""):
    """Search for the latest email with given subject and download its attachment"""

    if not is_outlook_running():
        return False, "Outlook is not running. Please start Outlook first."

    timeout = time.time() + (wait_minutes * 60)

    COMContext.initialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        try:
            namespace.Logon("", "", False, True)
        except:
            pass

        inbox = get_target_inbox(namespace, account_name)
        if inbox is None:
            return False, f"Cannot find an account matching '{account_name}' in Outlook"

        status_text = st.empty()
        progress_bar = st.progress(0)
        downloaded_files = []

        while time.time() < timeout:
            try:
                messages = inbox.Items
                messages.Sort("[ReceivedTime]", True)

                max_check = min(301, messages.Count + 1)
                for i in range(1, max_check):
                    try:
                        message = messages.Item(i)
                        subject = str(message.Subject)

                        if subject_keyword.lower() in subject.lower():
                            if message.Attachments.Count > 0:
                                for j in range(1, message.Attachments.Count + 1):
                                    attachment = message.Attachments.Item(j)
                                    filename = attachment.FileName
                                    filepath = DOWNLOAD_DIR / filename
                                    attachment.SaveAsFile(str(filepath))
                                    downloaded_files.append(filepath)

                                progress_bar.empty()
                                status_text.empty()
                                return True, downloaded_files
                            else:
                                progress_bar.empty()
                                status_text.empty()
                                return False, "Email found but no attachments"
                    except:
                        continue

                elapsed = int(time.time() - (timeout - wait_minutes * 60))
                total_wait = wait_minutes * 60
                progress_bar.progress(min(elapsed / total_wait, 1.0))
                status_text.text(f"⏳ Searching for '{subject_keyword}'... {elapsed}s / {total_wait}s")

                time.sleep(3)

            except Exception as e:
                time.sleep(5)

        progress_bar.empty()
        status_text.empty()
        return False, f"Timeout reached. No email found with subject: '{subject_keyword}'"

    except Exception as e:
        return False, f"Error: {str(e)}"
    finally:
        COMContext.uninitialize()


# =========================================================
# FUNCTION TO GET THE LATEST EXCEL FILE
# =========================================================
def get_latest_excel_file():
    """Get the latest source Excel file from the download directory (excludes our own output)"""
    excel_files = [
        f
        for f in list(DOWNLOAD_DIR.glob("*.xlsx")) + list(DOWNLOAD_DIR.glob("*.xls"))
        if f.stem.lower() != Path(CLEANED_FILE_NAME).stem.lower() and not f.name.startswith("~$")
    ]

    if not excel_files:
        return None

    return max(excel_files, key=lambda f: f.stat().st_mtime)


# =========================================================
# HOLIDAY LIST FOR L2 DUE DATE (WORKDAY.INTL) CALCULATIONS
# =========================================================
# As given verbatim. A block of entries between 25-Dec-18 and 21-Oct-19 is
# missing its year in the source list (a formatting glitch, not intentional
# — every other entry has one) — inferred as 2019 in _parse_holiday_date()
# based on chronological position, not silently: flagged here so this gets
# checked against the real calendar if the list is ever regenerated.
HOLIDAY_DATES_RAW = """
15-Mar-06
14-Apr-06
09-Aug-06
15-Aug-06
02-Oct-06
19-Oct-06
23-Oct-06
25-Dec-06
01-Jan-07
26-Jan-07
04-Mar-07
06-Apr-07
15-Aug-07
28-Aug-07
02-Oct-07
21-Oct-07
09-Nov-07
25-Dec-07
01-Jan-08
26-Jan-08
21-Mar-08
22-Mar-08
15-Aug-08
02-Oct-08
09-Oct-08
28-Oct-08
25-Dec-08
01-Jan-09
26-Jan-09
11-Mar-09
10-Apr-09
05-Aug-09
15-Aug-09
28-Sep-09
02-Oct-09
13-Oct-09
17-Oct-09
25-Dec-09
01-Jan-10
26-Jan-10
01-Mar-10
02-Apr-10
06-Jul-10
15-Aug-10
24-Aug-10
02-Oct-10
17-Oct-10
05-Nov-10
25-Dec-10
01-Jan-11
26-Jan-11
20-Mar-11
22-Apr-11
17-May-11
13-Aug-11
15-Aug-11
02-Oct-11
06-Oct-11
26-Oct-11
27-Oct-11
25-Dec-11
26-Jan-12
11-Feb-12
20-Feb-12
08-Mar-12
10-Mar-12
24-Mar-12
06-Apr-12
14-Apr-12
12-May-12
09-Jun-12
02-Aug-12
15-Aug-12
02-Oct-12
24-Oct-12
13-Nov-12
25-Dec-12
01-Jan-13
26-Jan-13
27-Mar-13
15-Aug-13
20-Aug-13
02-Oct-13
14-Oct-13
03-Nov-13
25-Dec-13
01-Jan-14
26-Jan-14
17-Mar-14
10-Apr-14
15-Aug-14
02-Oct-14
03-Oct-14
15-Oct-14
23-Oct-14
24-Oct-14
25-Dec-14
01-Jan-15
26-Jan-15
06-Mar-15
03-Apr-15
15-Aug-15
02-Oct-15
22-Oct-15
11-Nov-15
25-Dec-15
01-Jan-16
26-Jan-16
24-Mar-16
25-Mar-16
15-Aug-16
18-Aug-16
02-Oct-16
11-Oct-16
30-Oct-16
31-Oct-16
25-Dec-16
01-Jan-17
26-Jan-17
13-Mar-17
07-Aug-17
15-Aug-17
30-Sep-17
02-Oct-17
19-Oct-17
20-Oct-17
25-Dec-17
01-Jan-18
26-Jan-18
02-Mar-18
15-Aug-18
02-Oct-18
19-Oct-18
07-Nov-18
25-Dec-18
01-Jan
26-Jan
21-Mar
19-Apr
15-Aug
02-Oct
08-Oct
27-Oct
28-Oct
25-Dec
21-Oct-19
01-Jan-20
26-Jan-20
10-Mar-20
10-Apr-20
25-May-20
15-Aug-20
03-Aug-20
02-Oct-20
25-Oct-20
14-Nov-20
16-Nov-20
25-Dec-20
01-Jan-21
26-Jan-21
29-Mar-21
02-Apr-21
21-Jul-21
15-Aug-21
22-Aug-21
02-Oct-21
15-Oct-21
04-Nov-21
05-Nov-21
25-Dec-21
01-Jan-22
26-Jan-22
18-Mar-22
15-Apr-22
03-May-22
15-Aug-22
02-Oct-22
05-Oct-22
24-Oct-22
25-Oct-22
25-Dec-22
26-Jan-23
08-Mar-23
30-Aug-23
15-Aug-23
02-Oct-23
24-Oct-23
13-Nov-23
25-Dec-23
01-Jan-24
26-Jan-24
25-Mar-24
15-Aug-24
19-Aug-24
02-Oct-24
31-Oct-24
01-Nov-24
25-Dec-24
01-Jan-25
26-Jan-25
14-Mar-25
15-Aug-25
02-Oct-25
20-Oct-25
21-Oct-25
25-Dec-25
01-Jan-26
26-Jan-26
04-Mar-26
15-Aug-26
28-Aug-26
02-Oct-26
20-Oct-26
09-Nov-26
10-Nov-26
25-Dec-26
""".strip().splitlines()

_HOLIDAY_YEAR_LESS_INFERRED_YEAR = 2019
HOLIDAY_LIST_SHEET_NAME = "Holiday List"


def _parse_holiday_date(raw):
    raw = raw.strip()
    for fmt in ("%d-%b-%y", "%d-%b-%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    try:
        return datetime.strptime(raw, "%d-%b").replace(
            year=_HOLIDAY_YEAR_LESS_INFERRED_YEAR
        ).date()
    except ValueError:
        return None


HOLIDAY_DATES = [d for d in (_parse_holiday_date(r) for r in HOLIDAY_DATES_RAW) if d]


def _write_holiday_list_sheet(wb):
    """Write the static holiday calendar into a 'Holiday List' sheet (column
    A, real date values) and define a workbook-scoped named range 'a' over
    it, so Pre_Offer_checkwise!N's and Z's WORKDAY.INTL formulas can
    reference it exactly as `WORKDAY.INTL(E{row},1,11,a)` /
    `WORKDAY.INTL(X{row},1,11,a)` — matching the literal formulas given, not
    a rewritten equivalent. Must run before wb.save() in both
    clean_latest_excel_file (N) and add_l2_summary_columns (Z) — the named
    range just needs to exist somewhere in the saved file by then, not
    necessarily before the formula text referencing it is written."""
    if HOLIDAY_LIST_SHEET_NAME in wb.sheetnames:
        del wb[HOLIDAY_LIST_SHEET_NAME]
    ws = wb.create_sheet(HOLIDAY_LIST_SHEET_NAME)

    ws["A1"] = "Date"
    for i, d in enumerate(HOLIDAY_DATES, start=2):
        cell = ws[f"A{i}"]
        cell.value = d
        cell.number_format = "dd-mmm-yy"

    last_row = len(HOLIDAY_DATES) + 1
    wb.defined_names["a"] = DefinedName(
        "a", attr_text=f"'{HOLIDAY_LIST_SHEET_NAME}'!$A$2:$A${last_row}"
    )


# =========================================================
# FUNCTION TO CLEAN THE LATEST EXCEL FILE
# =========================================================
def clean_latest_excel_file():
    """Delete columns A to R from the latest Excel file"""
    try:
        latest_file = get_latest_excel_file()

        if latest_file is None:
            return False, "No Excel files found in the directory"

        wb = load_workbook(latest_file, data_only=True)

        sheet_name = next(
            (name for name in wb.sheetnames if "pre_offer_checkwise" in name.lower()),
            None,
        )
        if sheet_name is None:
            return False, "Sheet 'Pre_Offer_checkwise' not found in the latest file"

        ws = wb[sheet_name]

        # Delete every column matching one of these headers, wherever they sit
        headers_to_delete = {
            "national identity check (2)",
            "criminal records verification (2)",
            "criminal records verification (3)",
            "criminal records verification (4)",
            "criminal records verification (5)",
            "criminal records verification (6)",
            "criminal records verification (7)",
            "criminal records verification (8)",
            "criminal records verification (9)",
        }
        while True:
            header_col = next(
                (
                    cell.column
                    for cell in ws[2]
                    if str(cell.value or "").strip().lower() in headers_to_delete
                ),
                None,
            )
            if header_col is None:
                break
            ws.delete_cols(header_col)

        # Insert "Due Date" and "L1 TAT" columns at N and O, with formulas down
        # to the last row. N's WORKDAY.INTL calls reference the 'a' named range
        # (the holiday calendar written by _write_holiday_list_sheet below) —
        # order-independent since it's all just formula text until Excel opens
        # the saved file, but 'a' must exist by the time this function saves.
        ws.insert_cols(14, amount=2)
        ws["N2"] = "Due Date"
        ws["O2"] = "L1 TAT"
        ws["AC2"] = "Ars N01"
        for row in range(3, ws.max_row + 1):
            n_cell = ws[f"N{row}"]
            n_cell.value = (
                f'=IF(G{row}<>"",WORKDAY.INTL(G{row},1,11,a),WORKDAY.INTL(E{row},1,11,a))'
            )
            n_cell.number_format = "dd-mmm-yy"
            ws[f"O{row}"] = f'=IF(L{row}<=N{row},"IT","OT")'
            ws[f"AC{row}"] = f'="\'"&C{row}&"\'"&IF(C{row + 1}<>"",",","")'

        _write_holiday_list_sheet(wb)

        cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
        wb.save(str(cleaned_file_path))

        return True, [cleaned_file_path]

    except Exception as e:
        return False, f"Error cleaning file: {str(e)}"


# =========================================================
# FUNCTION TO BUILD THE ARS NUMBER LIST FROM THE CLEANED FILE
# =========================================================
def get_ars_number_list():
    """Read column C (ARS Number) of the cleaned tracker and build a quoted,
    comma-separated list, e.g. 'ARS1','ARS2','ARS3' — equivalent to the AC column formula."""
    cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
    if not cleaned_file_path.exists():
        return None, f"No {CLEANED_FILE_NAME} found. Run 'Clean Latest File' first."

    wb = load_workbook(cleaned_file_path, data_only=True)
    ws = wb["Pre_Offer_checkwise"]

    ars_numbers = [
        str(ws[f"C{row}"].value).strip()
        for row in range(3, ws.max_row + 1)
        if ws[f"C{row}"].value not in (None, "")
    ]
    if not ars_numbers:
        return None, "No ARS numbers found in column C"

    return ",".join(f"'{ars}'" for ars in ars_numbers), None


def _select_dropdown_option(page, name, option_matcher, timeout=30):
    """Select an option in a <select name=...> that Autobridge's MIS tool rebuilds via
    onchange handlers. Playwright locators re-resolve the DOM by selector on every call
    (no held element handle), so there's no Selenium-style staleness to guard against —
    this just polls the option text until the rebuild has produced a matching option.

    Verifies the selection actually stuck before returning, and retries the whole
    selection if not. This matters specifically for csv_query (Query): re-selecting
    access_time (Data Time Slab) — even to its already-selected value — silently
    resets csv_query back to blank shortly afterward, without rebuilding its option
    list. select_option() not raising an exception does NOT mean the selection
    survived; confirmed live that selecting csv_query immediately after access_time
    gets silently wiped, while doing the identical select_option() call after a
    pause sticks. Don't drop this verification thinking it's redundant — it's the
    fix for a real, reproduced site behavior, not defensive-programming padding."""
    selector = f"select[name='{name}']"
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            option_texts = page.locator(f"{selector} option").all_text_contents()
            match = next((t for t in option_texts if option_matcher(t)), None)
            if match is None:
                time.sleep(0.3)
                continue
            page.select_option(selector, label=match)
            time.sleep(0.5)
            current_text = page.eval_on_selector(
                selector,
                "el => el.options[el.selectedIndex] ? el.options[el.selectedIndex].text : null",
            )
            if current_text != match:
                continue
            return match
        except Exception:
            time.sleep(0.3)
    raise TimeoutError(f"Timed out waiting for a matching option in <select name='{name}'>")


def select_query(page, data_time_slab_keyword, query_name):
    """Change the Data Time Slab / Query selection on an already-logged-in MIS session."""
    _select_dropdown_option(
        page, "access_time", lambda t: data_time_slab_keyword.lower() in t.lower()
    )
    _select_dropdown_option(page, "csv_query", lambda t: t == query_name)


def fill_ars_number_field(page, ars_list, timeout=30):
    """Paste the ARS number list into the ARS No* field for the currently selected query.
    page.fill() re-locates by selector and auto-waits for the element to be actionable,
    which already covers the query tool rebuilding this field's containing table via JS
    whenever the Query dropdown changes — no manual staleness retry needed, unlike Selenium."""
    selector = "xpath=//*[@id='date']/tbody/tr[2]/td[2]/input"
    page.fill(selector, ars_list, timeout=timeout * 1000)


# =========================================================
# FUNCTION TO LOG INTO THE MIS EXPORT QUERY TOOL AND NAVIGATE TO A QUERY
# =========================================================
def run_ars_query(data_time_slab_keyword, query_name):
    """Log into the MIS export tool and select Host/Database/Data Time Slab/Query.
    Returns a Playwright Page with the browser left open on the query form.

    Uses playwright.sync_api.sync_playwright().start() rather than the
    `with sync_playwright() as p:` context-manager form deliberately: exiting
    that `with` block tears down the driver connection and closes the browser,
    but per an explicit prior user request the browser must stay open after
    this function returns (so it can be reviewed / reused for later queries
    on the same session) — so .stop()/.close() are never called here, mirroring
    the old Selenium code's "never call driver.quit()" behavior.

    Resets the asyncio event loop policy to WindowsProactorEventLoopPolicy
    before starting Playwright. Streamlit is built on Tornado, which forces
    the process-wide policy to WindowsSelectorEventLoopPolicy on Windows —
    but SelectorEventLoop can't spawn subprocesses on Windows, and Playwright's
    sync API needs to spawn its driver subprocess, so calling this from inside
    a running Streamlit app raises NotImplementedError from deep inside
    asyncio's subprocess machinery (confirmed by reproducing it directly: the
    exact same traceback occurs by just forcing WindowsSelectorEventLoopPolicy
    outside of Streamlit too). This is why the exact same code always worked
    when run as a standalone script (default policy is Proactor there) but
    always failed through the Streamlit UI."""
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    playwright = sync_playwright().start()
    # channel="chrome" launches the real installed Chrome binary rather than
    # Playwright's own bundled Chromium. This isn't cosmetic: on this network,
    # the bundled Chromium gets net::ERR_CONNECTION_CLOSED hitting the MIS
    # site (confirmed: curl and a real Chrome instance both reach it fine,
    # the bundled binary alone was blocked) — some corporate proxy/security
    # layer is evidently allowing the recognized Chrome executable through
    # while blocking the unrecognized one. Don't drop this thinking it's
    # unnecessary.
    browser = playwright.chromium.launch(
        headless=False, channel="chrome", args=["--start-maximized"]
    )
    context = browser.new_context(accept_downloads=True, no_viewport=True)
    page = context.new_page()
    page.set_default_timeout(30_000)
    # Login/query pages here are ordinary navigations, not downloads, so a
    # generous but bounded navigation timeout is safe (unlike Export's click,
    # handled separately in export_and_download via expect_download()).
    page.set_default_navigation_timeout(90_000)

    page.goto(MIS_LOGIN_URL)

    page.fill("xpath=//input[@type='text']", MIS_USERNAME)
    page.fill("xpath=//input[@type='password']", MIS_PASSWORD)
    page.click("xpath=//input[@value='Login']")

    _select_dropdown_option(page, "hostname", lambda t: t == MIS_HOST)
    _select_dropdown_option(page, "database", lambda t: t == MIS_DATABASE)
    select_query(page, data_time_slab_keyword, query_name)

    return page


# =========================================================
# FUNCTION TO CLICK EXPORT AND WAIT FOR THE DOWNLOAD TO COMPLETE
# =========================================================
def export_and_download(page, timeout=120):
    """Click the Export button and save the resulting download into MIS_EXPORT_DIR.
    Returns (file_path, error).

    Uses page.expect_download() rather than clicking then polling the filesystem
    (the old Selenium approach). That polling approach had a real bug: Selenium's
    .click() on a form-submit button blocks until the browser reports the
    resulting navigation complete, and a slow server-side report generation could
    hang that single .click() call itself for minutes — well before our own
    polling loop even started, since it only began counting after .click()
    returned. It surfaced as a raw, uncatchable "Read timed out (120)" from the
    underlying HTTP client. expect_download() sidesteps this entirely: it listens
    for the browser's actual download event directly, decoupled from whatever
    Chrome's navigation/page-load state is doing."""
    selector = "xpath=/html/body/div/div/div[2]/form/input"

    try:
        page.wait_for_selector(selector, timeout=30_000)
    except PlaywrightTimeoutError:
        # Distinguishes "never found the button" from "clicked it, no
        # download followed" below — both used to raise the exact same
        # PlaywrightTimeoutError from inside one shared try/except, making
        # "Export clicked but no download..." a misleading message when the
        # click never actually happened.
        return None, "Could not find the Export button within 30s"

    try:
        with page.expect_download(timeout=timeout * 1000) as download_info:
            page.click(selector)
        download = download_info.value
    except PlaywrightTimeoutError:
        return None, "Export clicked but no download completed within the timeout"

    dest_path = MIS_EXPORT_DIR / download.suggested_filename
    if dest_path.exists():
        # Repeated exports in the same session (e.g. re-running a query) can
        # reuse the same suggested filename — de-duplicate like Chrome does
        # for plain downloads, rather than overwriting the earlier file.
        stem, suffix = dest_path.stem, dest_path.suffix
        counter = 1
        while dest_path.exists():
            dest_path = MIS_EXPORT_DIR / f"{stem} ({counter}){suffix}"
            counter += 1
    download.save_as(str(dest_path))

    return dest_path, None


# =========================================================
# FUNCTION TO READ THE EXPORTED CSV (PLAIN OR ZIPPED)
# =========================================================
def _read_export_csv_rows(export_file_path):
    """Return all rows (including header) from the exported Advance Tracker file,
    whether it downloaded as a plain .csv or a .csv.zip."""
    if export_file_path.suffix.lower() == ".zip":
        with zipfile.ZipFile(export_file_path) as zf:
            csv_name = next(name for name in zf.namelist() if name.lower().endswith(".csv"))
            with zf.open(csv_name) as f:
                text = f.read().decode("utf-8", errors="replace")
    else:
        text = export_file_path.read_text(encoding="utf-8", errors="replace")

    return list(csv.reader(io.StringIO(text)))


# =========================================================
# FUNCTION TO FILTER/AGGREGATE DISCREPANCY ROWS INTO A "red remarks" TAB
# =========================================================
RED_REMARKS_SEVERITIES = {"Major Discrepancy", "Amber", "Minor Discrepancy"}
RED_REMARKS_CHECK_NAMES = {
    "UAN Check for Undisclosed Employment",
    "Dual Employment Verification via Form 26AS",
    "Criminal Records Verification",
    "National Identity Check",
    "India Court Record Check through Law Firm",
}


def add_red_remarks_sheet(export_file_path):
    """Filter the exported Advance Tracker data to check_severity in
    RED_REMARKS_SEVERITIES AND Check_unique_name in RED_REMARKS_CHECK_NAMES,
    then group by case_ars_no and join each ARS's closure_comments with ", "
    into a single row per unique ARS. Writes case_ars_no (column B) ->
    combined closure_comments (column C) into a new 'red remarks' sheet in
    the cleaned tracker — verified against a real export to reproduce the
    expected combined-comment text exactly (e.g. ARS 3055-016865)."""
    cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
    if not cleaned_file_path.exists():
        return None, f"{CLEANED_FILE_NAME} not found. Run 'Clean Latest File' first."

    rows = _read_export_csv_rows(export_file_path)
    header, data_rows = rows[0], rows[1:]

    required = ["check_severity", "Check_unique_name", "case_ars_no", "closure_comments"]
    missing = [c for c in required if c not in header]
    if missing:
        return None, f"Column(s) {', '.join(missing)} not found in the exported file"

    sev_col = header.index("check_severity")
    name_col = header.index("Check_unique_name")
    ars_col = header.index("case_ars_no")
    comment_col = header.index("closure_comments")

    filtered_rows = [
        row
        for row in data_rows
        if len(row) > comment_col
        and row[sev_col] in RED_REMARKS_SEVERITIES
        and row[name_col] in RED_REMARKS_CHECK_NAMES
    ]

    comments_by_ars = {}
    for row in filtered_rows:
        if not row[comment_col]:
            continue
        comments_by_ars.setdefault(row[ars_col], []).append(row[comment_col])

    wb = load_workbook(cleaned_file_path)
    sheet_name = "red remarks"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    ws["B1"] = "case_ars_no"
    ws["C1"] = "closure_comments"
    for i, (ars, comments) in enumerate(comments_by_ars.items(), start=2):
        ws[f"B{i}"] = ars
        ws[f"C{i}"] = ", ".join(comments)

    tracker_ws = wb["Pre_Offer_checkwise"]
    for row in range(3, tracker_ws.max_row + 1):
        tracker_ws[f"W{row}"] = (
            f"=IFERROR(VLOOKUP(C{row},'red remarks'!B:AB,2,0),\"\")"
        )

    wb.save(cleaned_file_path)

    return len(comments_by_ars), None


# =========================================================
# FUNCTION TO FILTER/SORT CASE HISTORY DATA INTO A "Form submisison - L2" TAB
# =========================================================
FORM_SUBMISSION_L2_COMMENT = (
    "Check moved to WIP. Antecedents populated from iBridge candidate submission."
)


def add_form_submission_l2_sheet(export_file_path):
    """Filter the exported Case History data to ACTION_COMMENTS matching the
    L2 form-submission comment, sort by ACTION_TAKEN_ON oldest-first, and copy
    the full matching rows into a 'Form submisison - L2' sheet in the cleaned tracker."""
    cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
    if not cleaned_file_path.exists():
        return None, f"{CLEANED_FILE_NAME} not found. Run 'Clean Latest File' first."

    rows = _read_export_csv_rows(export_file_path)
    header, data_rows = rows[0], rows[1:]

    if "ACTION_COMMENTS" not in header or "ACTION_TAKEN_ON" not in header:
        return None, "Column 'ACTION_COMMENTS' or 'ACTION_TAKEN_ON' not found in the exported file"

    comments_col = header.index("ACTION_COMMENTS")
    taken_on_col = header.index("ACTION_TAKEN_ON")

    filtered_rows = [
        row
        for row in data_rows
        if len(row) > comments_col and row[comments_col] == FORM_SUBMISSION_L2_COMMENT
    ]
    filtered_rows.sort(key=lambda row: row[taken_on_col] if len(row) > taken_on_col else "")

    wb = load_workbook(cleaned_file_path)
    sheet_name = "Form submisison - L2"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    ws.append(header)
    for row in filtered_rows:
        ws.append(row)

    tracker_ws = wb["Pre_Offer_checkwise"]
    tracker_ws["X2"] = "Form submisison - L2"
    for row in range(3, tracker_ws.max_row + 1):
        tracker_ws[f"X{row}"] = (
            f"=IFERROR(VLOOKUP(C{row},'Form submisison - L2'!B:T,12,0),\"\")"
        )

    wb.save(cleaned_file_path)

    return len(filtered_rows), None


# =========================================================
# FUNCTION TO FILTER/SORT CASE HISTORY DATA INTO A "L2 check addition (Oldest)" TAB
# =========================================================
CASE_REOPENED_ACTION = "New Status - Case Reopened"


def add_l2_check_addition_oldest_sheet(export_file_path):
    """Filter the exported Case History data to ACTION_TAKEN == 'New Status - Case
    Reopened', sort by ACTION_TAKEN_ON oldest-first, and copy the full matching
    rows into a 'L2 check addition (Oldest)' sheet in the cleaned tracker."""
    cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
    if not cleaned_file_path.exists():
        return None, f"{CLEANED_FILE_NAME} not found. Run 'Clean Latest File' first."

    rows = _read_export_csv_rows(export_file_path)
    header, data_rows = rows[0], rows[1:]

    if "ACTION_TAKEN" not in header or "ACTION_TAKEN_ON" not in header:
        return None, "Column 'ACTION_TAKEN' or 'ACTION_TAKEN_ON' not found in the exported file"

    action_col = header.index("ACTION_TAKEN")
    taken_on_col = header.index("ACTION_TAKEN_ON")

    filtered_rows = [
        row for row in data_rows if len(row) > action_col and row[action_col] == CASE_REOPENED_ACTION
    ]
    filtered_rows.sort(key=lambda row: row[taken_on_col] if len(row) > taken_on_col else "")

    wb = load_workbook(cleaned_file_path)
    sheet_name = "L2 check addition (Oldest)"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    ws.append(header)
    for row in filtered_rows:
        ws.append(row)

    tracker_ws = wb["Pre_Offer_checkwise"]
    tracker_ws["Y2"] = "L2 check addition (Oldest)"
    for row in range(3, tracker_ws.max_row + 1):
        tracker_ws[f"Y{row}"] = (
            f"=IFERROR(VLOOKUP(C{row},'L2 check addition (Oldest)'!B:T,12,0),\"\")"
        )

    wb.save(cleaned_file_path)

    return len(filtered_rows), None


# =========================================================
# FUNCTION TO FILTER/SORT "Sent Cases" DATA AND MAP INTO "L2 Report sent date" /
# "L2 Report sent severity" COLUMNS
# =========================================================
REPORT_SENT_STATUS = "Sent"
REPORT_SENT_TYPE = "Additional"


def _find_header_col_letter(ws, header_name, header_row=2):
    """Find the column letter of a header in a Pre_Offer_checkwise-style sheet
    (header row 2, per the row-1-is-a-title-line gotcha)."""
    for cell in ws[header_row]:
        if str(cell.value or "").strip().lower() == header_name.strip().lower():
            return cell.column_letter
    return None


def add_l2_report_sent_sheet(export_file_path):
    """Filter the exported Sent Cases data to Report Status == 'Sent' and
    Report Type == 'Additional', sort by report_sent_on oldest-first, and copy
    the full matching rows into a 'L2 Report Sent' sheet in the cleaned tracker.
    Overwrites the existing 'L2 Report sent date' / 'L2 Report sent severity'
    values in Pre_Offer_checkwise (every row 3..max_row gets a fresh formula
    below), mapping case_ars_no -> ARS Number (column C) to refill them
    (date-only, dd-mmm-yy, for the date column).

    Uses INDEX/MATCH rather than VLOOKUP for the mapping formulas: unlike the
    other export-derived sheets here, the relative position of case_ars_no vs.
    report_sent_on/report_severity in the Sent Cases export hasn't been verified
    against a live file, and VLOOKUP can't look left of its lookup column."""
    cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
    if not cleaned_file_path.exists():
        return None, f"{CLEANED_FILE_NAME} not found. Run 'Clean Latest File' first."

    rows = _read_export_csv_rows(export_file_path)
    header, data_rows = rows[0], rows[1:]

    required = ["Report Status", "Report Type", "report_sent_on", "case_ars_no", "report_severity"]
    missing = [c for c in required if c not in header]
    if missing:
        return None, f"Column(s) {', '.join(missing)} not found in the exported file"

    status_col = header.index("Report Status")
    type_col = header.index("Report Type")
    sent_on_col = header.index("report_sent_on")
    ars_col = header.index("case_ars_no")
    severity_col = header.index("report_severity")

    filtered_rows = [
        row
        for row in data_rows
        if len(row) > max(status_col, type_col)
        and row[status_col].strip().lower() == REPORT_SENT_STATUS.lower()
        and row[type_col].strip().lower() == REPORT_SENT_TYPE.lower()
    ]
    filtered_rows.sort(key=lambda row: row[sent_on_col] if len(row) > sent_on_col else "")

    wb = load_workbook(cleaned_file_path)
    sheet_name = "L2 Report Sent"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    ws.append(header)
    for row in filtered_rows:
        ws.append(row)

    # Reparse report_sent_on into a real date (date-only) in the audit sheet so
    # the INDEX/MATCH pull-through below carries an actual date value, not text
    # — the exported value is a full datetime string, and the destination's
    # dd-mmm-yy number format only applies to real date-typed values.
    sent_on_letter = get_column_letter(sent_on_col + 1)
    for i, row in enumerate(filtered_rows, start=2):
        raw_value = row[sent_on_col] if len(row) > sent_on_col else ""
        parsed_date = None
        if raw_value:
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                try:
                    parsed_date = datetime.strptime(raw_value.strip(), fmt).date()
                    break
                except ValueError:
                    continue
        if parsed_date is not None:
            ws[f"{sent_on_letter}{i}"] = parsed_date

    ars_letter = get_column_letter(ars_col + 1)
    severity_letter = get_column_letter(severity_col + 1)

    tracker_ws = wb["Pre_Offer_checkwise"]
    date_dest_letter = _find_header_col_letter(tracker_ws, "L2 Report sent date")
    severity_dest_letter = _find_header_col_letter(tracker_ws, "L2 Report sent severity")

    if date_dest_letter is None:
        return None, "Column 'L2 Report sent date' not found in Pre_Offer_checkwise"
    if severity_dest_letter is None:
        return None, "Column 'L2 Report sent severity' not found in Pre_Offer_checkwise"

    # Every row 3..max_row gets a fresh formula below, which already overwrites
    # whatever was there before — so clearing to None first would be a no-op.
    for row in range(3, tracker_ws.max_row + 1):
        date_cell = tracker_ws[f"{date_dest_letter}{row}"]
        date_cell.value = (
            f"=IFERROR(INDEX('{sheet_name}'!{sent_on_letter}:{sent_on_letter},"
            f"MATCH(C{row},'{sheet_name}'!{ars_letter}:{ars_letter},0)),\"\")"
        )
        date_cell.number_format = "dd-mmm-yy"

        tracker_ws[f"{severity_dest_letter}{row}"] = (
            f"=IFERROR(INDEX('{sheet_name}'!{severity_letter}:{severity_letter},"
            f"MATCH(C{row},'{sheet_name}'!{ars_letter}:{ars_letter},0)),\"\")"
        )

    wb.save(cleaned_file_path)

    return len(filtered_rows), None


# =========================================================
# FUNCTION TO UPDATE "Case Status" WHEN "L2 Report sent date" IS POPULATED
# =========================================================
CASE_STATUS_TARGETS = {"Insufficient", "On Hold"}
CASE_STATUS_COMPLETED = "Completed"


def update_case_status_from_l2_report_sent():
    """For Pre_Offer_checkwise rows whose Case Status (K) is currently
    'Insufficient' or 'On Hold', overwrite it to 'Completed' if that row's
    ARS number has a non-blank L2 Report sent date; otherwise leave it as-is.

    'Non-blank L2 Report sent date' is determined from the 'L2 Report Sent'
    sheet's own case_ars_no column (the same source add_l2_report_sent_sheet
    used to build Pre_Offer_checkwise!U's INDEX/MATCH formula) rather than by
    reading U's cell value directly — openpyxl never evaluates formulas, so
    U would read back as unevaluated/None regardless of what it resolves to
    once opened in Excel. This is a one-time value overwrite (not a formula):
    'leave unchanged' only makes sense as a snapshot of whatever Case Status
    currently holds, and a formula can't reference its own cell's prior
    value. Naturally idempotent across repeat runs — once a row is flipped to
    'Completed' it no longer matches CASE_STATUS_TARGETS, so later runs (even
    if the L2 Report Sent data changes) never touch it again."""
    cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
    if not cleaned_file_path.exists():
        return None, f"{CLEANED_FILE_NAME} not found. Run 'Clean Latest File' first."

    wb = load_workbook(cleaned_file_path)
    if "L2 Report Sent" not in wb.sheetnames:
        return None, "'L2 Report Sent' sheet not found. Run the Sent Cases query first."

    l2_ws = wb["L2 Report Sent"]
    # case_ars_no's column position in this sheet follows the raw Sent Cases
    # CSV's own column order (it was appended as-is, not reordered to a fixed
    # layout), so it must be looked up by header name rather than assumed —
    # confirmed it sits at column F, not B, against a real export.
    l2_ars_letter = _find_header_col_letter(l2_ws, "case_ars_no", header_row=1)
    if l2_ars_letter is None:
        return None, "Column 'case_ars_no' not found in 'L2 Report Sent' sheet"

    ars_with_report_sent = {
        l2_ws[f"{l2_ars_letter}{row}"].value
        for row in range(2, l2_ws.max_row + 1)
        if l2_ws[f"{l2_ars_letter}{row}"].value not in (None, "")
    }

    tracker_ws = wb["Pre_Offer_checkwise"]
    status_letter = _find_header_col_letter(tracker_ws, "Case Status")
    if status_letter is None:
        return None, "Column 'Case Status' not found in Pre_Offer_checkwise"

    updated_count = 0
    for row in range(3, tracker_ws.max_row + 1):
        status_cell = tracker_ws[f"{status_letter}{row}"]
        if status_cell.value not in CASE_STATUS_TARGETS:
            continue
        ars = tracker_ws[f"C{row}"].value
        if ars in ars_with_report_sent:
            status_cell.value = CASE_STATUS_COMPLETED
            updated_count += 1

    wb.save(cleaned_file_path)

    return updated_count, None


# =========================================================
# FUNCTION TO ADD L2 DUE DATE / L2 TAT / L2 CHECK STATUS COLUMNS
# =========================================================
def add_l2_summary_columns():
    """Add 'L2 Due Date' (Z), 'L2 TAT' (AA), and 'L2 check status' (AB) columns to
    Pre_Offer_checkwise. Z is derived from X ('Form submisison - L2', added by
    add_form_submission_l2_sheet) via WORKDAY.INTL against the 'a' named range
    (the holiday calendar written by _write_holiday_list_sheet), and AA from
    Z and U ('L2 Report sent date'). AB checks X, Y ('L2 check addition
    (Oldest)', added by add_l2_check_addition_oldest_sheet — an independent
    lookup from X, not derived from it, unlike Z), U, and K ('Case Status').
    Must run after add_form_submission_l2_sheet (X), add_l2_check_addition_oldest_sheet
    (Y), and clean_latest_excel_file (the 'a' named range)."""
    cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
    if not cleaned_file_path.exists():
        return False, f"{CLEANED_FILE_NAME} not found. Run 'Clean Latest File' first."

    wb = load_workbook(cleaned_file_path)
    ws = wb["Pre_Offer_checkwise"]

    ws["Z2"] = "L2 Due Date"
    ws["AA2"] = "L2 TAT"
    ws["AB2"] = "L2 check status"

    for row in range(3, ws.max_row + 1):
        z_cell = ws[f"Z{row}"]
        z_cell.value = f'=IF(X{row}="","",WORKDAY.INTL(X{row},1,11,a))'
        z_cell.number_format = "dd-mmm-yy"
        ws[f"AA{row}"] = (
            f'=IFERROR(IF(Z{row}="","",IF(INT(U{row})<=INT(Z{row}),"IT","OT")),"")'
        )
        ws[f"AB{row}"] = (
            f'=IF(AND(X{row}<>"",Y{row}<>"",U{row}<>""),K{row},'
            f'IF(AND(X{row}="",Y{row}<>""),"Pending at candidate",'
            f'IF(X{row}<>"","Work In Progress","")))'
        )

    wb.save(cleaned_file_path)

    return True, None


# =========================================================
# FUNCTION TO BUILD THE "Status view" PIVOT-TABLE REPORT SHEET
# =========================================================
# Modeled on the "Status view" tab of the reference
# "Mphasis_Limited_Pre_Offer_checkwise Dashboard V1.1.xlsx" workbook, which
# holds several live Excel PivotTables (Received cases, L1/L2 report
# severity, L1/L2 TAT, L2 check status, Case status). That reference file is
# NOT opened/linked anywhere in this app — only its layout/shape was used as
# a guide.
#
# openpyxl cannot create real PivotTable objects, so this drives a hidden
# Excel instance via COM automation (win32com — the same technique already
# used for Outlook) to build genuine, live, refreshable PivotTables straight
# off the final output workbook (CLEANED_FILE_NAME). Verified against the
# real file: the resulting
# PivotTables survive apply_professional_styling()'s later openpyxl
# load+save untouched and still refresh correctly afterward.
STATUS_VIEW_SHEET_NAME = "Status view"

# Late-bound win32com automation has no named Excel constants available, so
# these are hardcoded from the stable, documented XlPivotFieldOrientation /
# XlConsolidationFunction / XlPivotTableSourceType / xlUp / xlToLeft enums.
_XL_UP = -4162
_XL_TO_LEFT = -4159
_XL_DATABASE = 1
_XL_ROW_FIELD = 1
_XL_COLUMN_FIELD = 2
_XL_COUNT = -4112


def _com_last_row(ws, key_col):
    """1-based row of the last non-empty cell in key_col, scanning upward
    from the bottom of the sheet (mirrors Ctrl+Up)."""
    return ws.Cells(ws.Rows.Count, key_col).End(_XL_UP).Row


def _com_last_col(ws, header_row):
    """1-based column of the last non-empty header cell in header_row,
    scanning leftward from the right edge of the sheet (mirrors Ctrl+Left)."""
    return ws.Cells(header_row, ws.Columns.Count).End(_XL_TO_LEFT).Column


def _rgb_hex_to_ole_bgr(hex_color):
    """Excel COM color properties (Tab.Color, Interior.Color, Font.Color) take
    an OLE COLOR (0x00BBGGRR), not RGB hex — convert once here rather than
    hand-computing the byte order at each call site."""
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return (b << 16) | (g << 8) | r


# Same navy/white theme apply_professional_styling() already uses for every
# other sheet's header row (HEADER_FILL/HEADER_FONT, defined further below) —
# duplicated as hex constants here since those Font/PatternFill objects are
# openpyxl types, not usable against COM Range objects.
_STATUS_VIEW_TITLE_FILL = _rgb_hex_to_ole_bgr("1F4E78")
_STATUS_VIEW_TITLE_FONT_COLOR = _rgb_hex_to_ole_bgr("FFFFFF")
_STATUS_VIEW_TAB_COLOR = _rgb_hex_to_ole_bgr("2E7D32")
_PIVOT_TABLE_STYLE = "PivotStyleMedium9"


def _hide_blank_pivot_item(pivot_field):
    """Hide a PivotField's blank/no-value category. Its PivotItem.Name is
    sometimes the literal empty string and sometimes the literal text
    '(blank)' depending on the field — confirmed inconsistent across fields
    in a live run — so both are checked. No-ops if there's no blank item."""
    items = pivot_field.PivotItems()
    for i in range(1, items.Count + 1):
        item = items.Item(i)
        if item.Name in ("", "(blank)"):
            item.Visible = False
            return


def _add_pivot_table(status_ws, cache, anchor_col, table_name, title, row_field, data_field,
                      col_field=None, hide_blank_row=False, hide_blank_col=False,
                      extra_row_fields=None, row_number_format=None):
    """Create one live PivotTable anchored at (row 3, anchor_col), with a
    bold plain-text title one row above it (row 2) — mirroring the reference
    dashboard's layout of a label sitting just above each pivot.

    extra_row_fields, if given, is a list of (field_name, number_format)
    tuples added as OUTER row fields before row_field, in the given order —
    for a caller that needs an explicit multi-level row hierarchy in place of
    Excel's native date Group() feature (see add_status_view_sheet for why
    that's broken via this COM path). No current pivot uses this.

    Returns the column immediately to the right of this table's actual
    footprint (plus a 1-column gap), so the next pivot table can be anchored
    there without ever overlapping this one — column widths here depend on
    how many distinct categories are actually in this run's data, so this is
    computed, not a fixed offset."""
    status_ws.Cells(2, anchor_col).Value = title

    pt = cache.CreatePivotTable(TableDestination=status_ws.Cells(3, anchor_col), TableName=table_name)
    try:
        pt.TableStyle2 = _PIVOT_TABLE_STYLE
    except Exception:
        pass

    # Field manipulation beyond Orientation/Position — NumberFormat, hiding a
    # blank item — reliably fails here unless it happens AFTER the table has
    # a data field and has been refreshed at least once (confirmed live: the
    # exact same NumberFormat assignment that fails right after Orientation
    # succeeds once issued post-refresh instead). Set every field's
    # Orientation/Position first, then AddDataField+Refresh, then apply
    # NumberFormat/blank-hiding as a second pass.
    position = 1
    extra_pfs = []
    for field_name, number_format in extra_row_fields or []:
        pf = pt.PivotFields(field_name)
        pf.Orientation = _XL_ROW_FIELD
        pf.Position = position
        extra_pfs.append((pf, number_format))
        position += 1

    row_pf = pt.PivotFields(row_field)
    row_pf.Orientation = _XL_ROW_FIELD
    row_pf.Position = position

    col_pf = None
    if col_field:
        col_pf = pt.PivotFields(col_field)
        col_pf.Orientation = _XL_COLUMN_FIELD

    pt.AddDataField(pt.PivotFields(data_field), f"Count of {data_field}", _XL_COUNT)
    pt.RefreshTable()

    for pf, number_format in extra_pfs:
        if number_format:
            try:
                pf.NumberFormat = number_format
            except Exception:
                pass
    if row_number_format:
        try:
            row_pf.NumberFormat = row_number_format
        except Exception:
            pass
    if hide_blank_row:
        _hide_blank_pivot_item(row_pf)
    if hide_blank_col and col_pf is not None:
        _hide_blank_pivot_item(col_pf)
    pt.RefreshTable()

    right_col = pt.TableRange1.Column + pt.TableRange1.Columns.Count - 1

    # Navy/white banner across the pivot's actual width, matching the
    # navy-header theme apply_professional_styling() uses on every other
    # sheet — only known now that the table's real width has been measured.
    title_range = status_ws.Range(status_ws.Cells(2, anchor_col), status_ws.Cells(2, right_col))
    title_range.Merge()
    title_range.Interior.Color = _STATUS_VIEW_TITLE_FILL
    title_range.Font.Color = _STATUS_VIEW_TITLE_FONT_COLOR
    title_range.Font.Bold = True
    title_range.Font.Size = 11
    title_range.HorizontalAlignment = -4108  # xlCenter
    status_ws.Rows(2).RowHeight = 20

    return right_col + 2


def add_status_view_sheet():
    """Build a 'Status view' sheet in the final output workbook holding 7 genuine,
    live Excel PivotTables: Received cases, Case status, L1 report severity,
    L1 TAT, L2 TAT (by Case status), L2 check status, and L2 report-sent
    severity — mirroring the reference dashboard's pivots, computed fresh
    each run, not linked to that file. All 7 are built off the single
    Pre_Offer_checkwise-sourced `cache` (no pivot here reads a separate audit
    sheet as its own PivotCache source).

    Runs via a hidden Excel COM automation session (see the constants/helpers
    above) rather than openpyxl, because openpyxl cannot create real
    PivotTable objects. This also sidesteps openpyxl's other limitation for
    this feature: Pre_Offer_checkwise's N/O/U/V/W/X/Y/Z/AA/AB columns are all
    formulas that openpyxl never evaluates (see clean_latest_excel_file /
    add_l2_summary_columns) — opening the file in real Excel and forcing a
    full recalculation resolves them for real, so the pivots can aggregate
    directly on Pre_Offer_checkwise's own columns instead of a Python
    reimplementation of that formula logic.

    Gotcha (still relevant to why every date row field here — Case Received
    Date, L1 Report Sent Date, L2 Report sent date — is a plain row field
    with a row_number_format rather than a nested hierarchy): Excel's native
    date-grouping (Range.Group, e.g. for a Year>Month>Day hierarchy) reliably
    FAILED via this COM automation path in testing — reproduced even on a
    trivial, freshly hand-built numeric field with no connection to this
    app's data, so it isn't a data-quality issue, just something about
    Group() via this binding/environment. _add_pivot_table's
    extra_row_fields param exists for a caller that needs an explicit
    multi-level row hierarchy in place of Group() (PivotField.Position
    controls nesting order explicitly, since assignment order does NOT
    determine it — confirmed: fields nest by their underlying column order in
    the source range unless Position is set), but no current pivot needs
    it."""
    cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
    if not cleaned_file_path.exists():
        return False, f"{CLEANED_FILE_NAME} not found. Run 'Clean Latest File' first."

    COMContext.initialize()
    excel = None
    excel_pid = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        # Confirmed live: Quit() alone unreliably leaves EXCEL.EXE resident
        # when Visible=False (reproduced repeatedly, including with zero
        # workbooks open at Quit() time) — a known limitation of hidden Excel
        # COM automation, not a bug in the cleanup logic below. The PID is
        # captured up front so the finally block can force-kill it if Quit()
        # doesn't actually end the process.
        excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(str(cleaned_file_path))
        excel.CalculateFullRebuild()

        tracker_ws = wb.Worksheets("Pre_Offer_checkwise")
        last_row = _com_last_row(tracker_ws, 3)  # column C = ARS Number, always populated
        last_col = _com_last_col(tracker_ws, header_row=2)
        source_range = tracker_ws.Range(tracker_ws.Cells(2, 1), tracker_ws.Cells(last_row, last_col))

        for ws in list(wb.Worksheets):
            if ws.Name == STATUS_VIEW_SHEET_NAME:
                ws.Delete()
        status_ws = wb.Worksheets.Add(Before=wb.Worksheets(1))
        status_ws.Name = STATUS_VIEW_SHEET_NAME

        cache = wb.PivotCaches().Create(SourceType=_XL_DATABASE, SourceData=source_range)

        next_col = 2
        next_col = _add_pivot_table(
            status_ws, cache, next_col, "PT_ReceivedCases", "Received cases",
            row_field="Case Received Date", data_field="ARS Number",
            row_number_format="dd-mmm-yy",
        )
        next_col = _add_pivot_table(
            status_ws, cache, next_col, "PT_CaseStatus", "Case status",
            row_field="Case Status", data_field="ARS Number",
        )
        next_col = _add_pivot_table(
            status_ws, cache, next_col, "PT_L1Severity", "L1 reports sent severity",
            row_field="L1 Report Sent Date", col_field="L1 Report Severity", data_field="ARS Number",
            row_number_format="dd-mmm-yy", hide_blank_row=True, hide_blank_col=True,
        )
        next_col = _add_pivot_table(
            status_ws, cache, next_col, "PT_L1TAT", "L1 TAT",
            row_field="L1 Report Sent Date", col_field="L1 TAT", data_field="ARS Number",
            row_number_format="dd-mmm-yy", hide_blank_row=True,
        )
        # Pre_Offer_checkwise's L2 TAT / L2 check status columns are formulas
        # that look up 'Form submisison - L2' / 'L2 Report Sent' by sheet
        # name — if "Clean Latest File" has run but the MIS query chain
        # hasn't yet, those sheets don't exist and the formulas resolve to
        # real Excel errors, which breaks PivotFields() on those columns
        # entirely (confirmed live: "PivotFields method of PivotTable class
        # failed"). Skip these two pivots until both aux sheets exist; the
        # MIS button handler rebuilds this whole sheet again once they do.
        sheet_names = [ws.Name for ws in wb.Worksheets]
        if "Form submisison - L2" in sheet_names and "L2 Report Sent" in sheet_names:
            next_col = _add_pivot_table(
                status_ws, cache, next_col, "PT_L2TAT", "L2 reports TAT",
                row_field="Case Status", col_field="L2 TAT", data_field="ARS Number",
                hide_blank_col=True,
            )
            next_col = _add_pivot_table(
                status_ws, cache, next_col, "PT_L2CheckStatus", "L2 check status",
                row_field="L2 check status", data_field="ARS Number",
                hide_blank_row=True,
            )

        # Sourced directly from Pre_Offer_checkwise's own 'L2 Report sent
        # date' / 'L2 Report sent severity' columns (same `cache` as the
        # other 6 pivots) rather than the 'L2 Report Sent' audit sheet — those
        # two tracker columns exist in the source tracker from the start (see
        # update_case_status_from_l2_report_sent / add_l2_report_sent_sheet),
        # so this no longer needs to be gated on that sheet's presence, and
        # mirrors the 'L1 reports sent severity' pivot's own pattern (date
        # row field + severity column field + ARS Number count) instead of
        # the old separate-cache/Year-Month-helper-column workaround.
        next_col = _add_pivot_table(
            status_ws, cache, next_col, "PT_L2SentSeverity", "L2 reports sent severity",
            row_field="L2 Report sent date", col_field="L2 Report sent severity", data_field="ARS Number",
            row_number_format="dd-mmm-yy", hide_blank_row=True, hide_blank_col=True,
        )

        status_ws.Tab.Color = _STATUS_VIEW_TAB_COLOR
        status_ws.UsedRange.Columns.AutoFit()

        wb.Save()
        return True, None

    except Exception as e:
        return False, f"Error building Status view pivots: {str(e) or type(e).__name__}"

    finally:
        # Release COM references before Quit(), and before the PID-based
        # fallback kill below — matters for the fallback to reliably observe
        # process exit (see excel_pid comment above).
        tracker_ws = source_range = status_ws = cache = None
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        wb = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        excel = None
        COMContext.uninitialize()

        if excel_pid:
            time.sleep(2)
            still_running = subprocess.run(
                ["tasklist", "/FI", f"PID eq {excel_pid}"], capture_output=True, text=True
            )
            if str(excel_pid) in still_running.stdout:
                subprocess.run(["taskkill", "/PID", str(excel_pid), "/F"], capture_output=True, text=True)


# =========================================================
# FUNCTION TO APPLY PROFESSIONAL HEADER STYLING ACROSS ALL SHEETS
# =========================================================
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
HEADER_BORDER = Border(bottom=Side(style="thin", color="9CA3AF"))
AUDIT_SHEET_TAB_COLOR = "2E75B6"
AUDIT_SHEET_NAMES = [
    "red remarks",
    "Form submisison - L2",
    "L2 check addition (Oldest)",
    "L2 Report Sent",
    HOLIDAY_LIST_SHEET_NAME,
]


def _style_header_row(ws, header_row, min_col=1, max_col=None):
    """Bold white-on-navy header styling, a frozen pane below the header, and
    a readable minimum column width (capped so a long free-text header
    doesn't blow a column out to its full length). Purely cosmetic — only
    touches the header row's cell styling and column/row dimensions, never
    cell values or formulas."""
    max_col = max_col or ws.max_column
    for col in range(min_col, max_col + 1):
        cell = ws.cell(row=header_row, column=col)
        if cell.value in (None, ""):
            continue
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = HEADER_BORDER
        col_letter = get_column_letter(col)
        target_width = min(max(len(str(cell.value)) + 4, 14), 40)
        current_width = ws.column_dimensions[col_letter].width
        if not current_width or current_width < target_width:
            ws.column_dimensions[col_letter].width = target_width
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate
    ws.row_dimensions[header_row].height = 24


def apply_professional_styling():
    """Final cosmetic pass over the final output workbook: bold/colored header
    rows, frozen header panes, readable column widths, and tab colors across
    Pre_Offer_checkwise and every audit sheet this app creates, plus text
    wrapping for 'red remarks'!closure_comments (which can run to
    paragraph-length combined comments — see add_red_remarks_sheet), and
    hidden gridlines + a frozen pane on 'Status view' (its navy pivot-title
    banners and PivotTable styling are set directly via COM in
    add_status_view_sheet, since openpyxl can't touch PivotTable formatting —
    this only adds the plain sheet-level view properties on top). Purely
    visual, never touches cell values or formulas. Runs last, since it needs
    every sheet/header this pipeline creates (N/O/AC, W, X, Y, Z/AA/AB, U/L2
    Report sent severity, Status view) to already be in place."""
    cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
    if not cleaned_file_path.exists():
        return False, f"{CLEANED_FILE_NAME} not found. Run 'Clean Latest File' first."

    wb = load_workbook(cleaned_file_path)

    tracker_ws = wb["Pre_Offer_checkwise"]
    _style_header_row(tracker_ws, header_row=2)

    for sheet_name in AUDIT_SHEET_NAMES:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        _style_header_row(ws, header_row=1)
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = AUDIT_SHEET_TAB_COLOR

    if "red remarks" in wb.sheetnames:
        remarks_ws = wb["red remarks"]
        remarks_ws.column_dimensions["C"].width = 80
        wrap_top = Alignment(wrap_text=True, vertical="top")
        for row in range(2, remarks_ws.max_row + 1):
            remarks_ws[f"C{row}"].alignment = wrap_top

    if STATUS_VIEW_SHEET_NAME in wb.sheetnames:
        # Purely sheet-level view properties (gridlines, freeze panes) —
        # confirmed live that this kind of openpyxl load+save leaves the
        # sheet's real PivotTables intact and still refreshable (see
        # add_status_view_sheet's docstring for how that was verified).
        status_ws = wb[STATUS_VIEW_SHEET_NAME]
        status_ws.sheet_view.showGridLines = False
        status_ws.freeze_panes = "A5"

    wb.save(cleaned_file_path)

    return True, None


# =========================================================
# FUNCTION TO LIST FILES IN DOWNLOAD DIRECTORY
# =========================================================
def list_downloaded_files():
    """List all files in the download directory"""
    files = [f for f in DOWNLOAD_DIR.glob("*") if f.is_file()]
    if not files:
        st.caption("No files in download directory yet")
        return

    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)

    file_data = [
        {
            "File": f.name,
            "Size (KB)": f"{f.stat().st_size / 1024:.1f}",
            "Modified": datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
        }
        for f in files
    ]
    st.table(file_data)


# =========================================================
# MAIN APPLICATION
# =========================================================

with st.expander("Step 1: Manually trigger the tracker"):
    st.code(TRIGGER_URL, language="html")
    st.caption("Open this URL, then wait for the email to arrive in Outlook")

st.subheader("Step 2: Download & Clean")

mail_subject = st.text_input(
    "Email subject keyword",
    value="Mphasis_Limited | ARS Non Standard Tracker | Pre_Offer_checkwise",
)
account_name = st.text_input(
    "Outlook account (leave blank to use Outlook's default account)",
    value="jay.chaudhary@authbridge.com",
    help="If you have multiple accounts in Outlook, this must match the one whose "
    "Inbox actually receives the tracker email — GetDefaultFolder only searches "
    "the default account otherwise.",
)
wait_time = st.number_input("Search time (minutes)", min_value=1, max_value=15, value=5, step=1)

st.caption("✅ Outlook is running" if is_outlook_running() else "❌ Outlook is NOT running")

col1, col2, col3 = st.columns(3)
with col1:
    download_clicked = st.button("📥 Download Latest", type="primary", use_container_width=True)
with col2:
    clean_clicked = st.button("🗑️ Clean Latest File", use_container_width=True)
with col3:
    if st.button("📂 Open Folder", use_container_width=True):
        os.startfile(str(DOWNLOAD_DIR))


if download_clicked:
    if not mail_subject:
        st.warning("Please enter the email subject keyword")
    elif not is_outlook_running():
        st.error("Outlook is not running. Please start Outlook first.")
    else:
        with st.spinner("Searching Outlook..."):
            success, result = download_latest_from_email(mail_subject, wait_time, account_name)

        if success:
            st.success(f"Downloaded {len(result)} file(s)")
            for file in result:
                st.caption(f"{file.name} ({file.stat().st_size / 1024:.2f} KB)")
        else:
            st.error(result)


if clean_clicked:
    with st.spinner("Cleaning columns A through R..."):
        success, result = clean_latest_excel_file()

    if success:
        add_status_view_sheet()
        apply_professional_styling()
        st.success("File cleaned successfully")
        for file in result:
            st.caption(f"{file.name} ({file.stat().st_size / 1024:.2f} KB)")
    else:
        st.error(result)


st.markdown("---")
st.subheader("📁 Downloaded Files")
list_downloaded_files()


st.markdown("---")
st.subheader("Step 3: Run ARS Query in MIS")
st.caption(
    "Opens the MIS export tool, logs in, and runs Advance Tracker, Case History, "
    "and Sent Cases queries with the ARS numbers"
)

if st.button("🔎 Run ARS Query in MIS", type="primary", use_container_width=True):
    ars_list, error = get_ars_number_list()

    if error:
        st.error(error)
    else:
        with st.spinner("Logging into MIS and navigating to Advance Tracker..."):
            try:
                page = run_ars_query("Case Query", "Advance Tracker")
                fill_ars_number_field(page, ars_list)
                st.success(f"Pasted {ars_list.count(',') + 1} ARS number(s) into the query form")

                downloaded_file, download_error = export_and_download(page)
                if downloaded_file:
                    st.success(f"Exported: {downloaded_file.name}")

                    with st.spinner("Filtering 'Major Discrepancy' rows into 'red remarks' tab..."):
                        red_remarks_count, red_remarks_error = add_red_remarks_sheet(downloaded_file)

                    if red_remarks_error:
                        st.error(red_remarks_error)
                    else:
                        st.success(
                            f"Added 'red remarks' tab with {red_remarks_count} Major Discrepancy row(s)"
                        )
                else:
                    st.warning(download_error)

                with st.spinner("Navigating to Case History and pasting ARS numbers..."):
                    select_query(page, "Case Query", "Case History")
                    fill_ars_number_field(page, ars_list)

                case_history_file, case_history_error = export_and_download(page)
                if case_history_file:
                    st.success(f"Exported: {case_history_file.name}")

                    with st.spinner("Filtering & sorting into 'Form submisison - L2' tab..."):
                        l2_count, l2_error = add_form_submission_l2_sheet(case_history_file)

                    if l2_error:
                        st.error(l2_error)
                    else:
                        st.success(
                            f"Added 'Form submisison - L2' tab with {l2_count} row(s)"
                        )

                    with st.spinner("Filtering & sorting into 'L2 check addition (Oldest)' tab..."):
                        reopened_count, reopened_error = add_l2_check_addition_oldest_sheet(
                            case_history_file
                        )

                    if reopened_error:
                        st.error(reopened_error)
                    else:
                        st.success(
                            f"Added 'L2 check addition (Oldest)' tab with {reopened_count} row(s)"
                        )

                else:
                    st.warning(case_history_error)

                with st.spinner("Navigating to Sent Cases and pasting ARS numbers..."):
                    select_query(page, "Case Query", "Sent Cases")
                    fill_ars_number_field(page, ars_list)

                sent_cases_file, sent_cases_error = export_and_download(page)
                if sent_cases_file:
                    st.success(f"Exported: {sent_cases_file.name}")

                    with st.spinner(
                        "Filtering 'Sent' / 'Additional' rows and mapping into "
                        "'L2 Report sent date' / 'L2 Report sent severity'..."
                    ):
                        sent_count, sent_error = add_l2_report_sent_sheet(sent_cases_file)

                    if sent_error:
                        st.error(sent_error)
                    else:
                        st.success(
                            f"Added 'L2 Report Sent' tab with {sent_count} row(s) and mapped "
                            "'L2 Report sent date' / 'L2 Report sent severity'"
                        )
                else:
                    st.warning(sent_cases_error)

                with st.spinner(
                    "Marking 'Insufficient'/'On Hold' cases as 'Completed' where "
                    "L2 Report sent date is populated..."
                ):
                    status_count, status_error = update_case_status_from_l2_report_sent()

                if status_error:
                    st.error(status_error)
                else:
                    st.success(f"Updated Case Status to 'Completed' for {status_count} row(s)")

                with st.spinner("Adding L2 Due Date / L2 TAT / L2 check status columns..."):
                    summary_success, summary_error = add_l2_summary_columns()

                if summary_success:
                    st.success("Added 'L2 Due Date', 'L2 TAT', and 'L2 check status' columns")
                else:
                    st.error(summary_error)

                with st.spinner("Building 'Status view' summary report..."):
                    status_view_success, status_view_error = add_status_view_sheet()

                if status_view_success:
                    st.success("Added 'Status view' summary report")
                else:
                    st.error(status_view_error)

                with st.spinner("Applying professional formatting..."):
                    style_success, style_error = apply_professional_styling()

                if style_success:
                    st.success("Applied header styling across all tabs")
                else:
                    st.error(style_error)
            except Exception as e:
                st.error(f"Error running MIS query: {str(e) or type(e).__name__}")
                st.exception(e)


st.markdown("---")
st.subheader("Step 4: Download Final Report")

final_report_path = DOWNLOAD_DIR / CLEANED_FILE_NAME
if final_report_path.exists():
    st.download_button(
        "⬇️ Download Final Report",
        data=final_report_path.read_bytes(),
        file_name=CLEANED_FILE_NAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    st.caption(
        f"{CLEANED_FILE_NAME} ({final_report_path.stat().st_size / 1024:.1f} KB, "
        f"last updated {datetime.fromtimestamp(final_report_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')})"
    )
else:
    st.caption("No final report yet — run 'Clean Latest File' first.")
