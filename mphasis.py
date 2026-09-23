import streamlit as st
import time
import os
from decimal import Decimal
from pathlib import Path
from datetime import datetime, date
import subprocess
import threading
import win32com.client
import win32process
import pythoncom
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
import mysql.connector
from mysql.connector import Error as MySQLError
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
# CHECKPOINT_LIVE DATABASE CONFIG
# =========================================================
# Direct DB connection replaces the earlier Playwright/Selenium browser
# automation of the MIS export-query website (https://mis.authbridge.com/export_query/)
# for pulling Advance Tracker / Case History / Sent Cases data. The MIS site's
# "Host"/"Database" dropdowns were just display labels ("Bridge Live") over
# this same underlying checkpoint_live schema — this connects to it directly
# instead of driving the website's UI to run the same queries.
DB_HOST = os.getenv("DB_HOST")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_NAME = os.getenv("DB_NAME")
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")


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
def get_ars_numbers():
    """Read column C (ARS Number) of the cleaned tracker and return the raw list
    of ARS number strings (used as bind parameters for the DB queries below —
    no manual quoting needed, mysql-connector parameterizes the IN clause)."""
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

    return ars_numbers, None


# =========================================================
# DATABASE QUERY LAYER — REPLACES THE PLAYWRIGHT/MIS-WEBSITE AUTOMATION
# =========================================================
# The three queries below are exactly what the MIS export-query website ran
# server-side for "Advance Tracker" / "Case History" / "Sent Cases" (its
# Host/Database dropdowns were just display labels — "Bridge Live" — over
# this same checkpoint_live schema). Querying the DB directly removes the
# whole browser-automation layer (login, dropdown selection, Export-click,
# CSV/zip download+parse) — same data, same column names, no browser.
#
# `check_id1` is the literal placeholder text from the site's own SQL (its
# ARS-number-list substitution point) — kept verbatim rather than rewritten,
# and swapped for a parameterized `%s,%s,...` IN-list at execute time via
# _execute_ars_query() so ARS numbers are always bound as query parameters,
# never string-interpolated into the SQL text.
ADVANCE_TRACKER_QUERY = """
SELECT ecc.Case_Check_id,case_ars_no,emc.Company_name,
CONCAT(first_name,' ',IFNULL(middle_name,''),' ',IFNULL(last_name,'')) AS Candidate_name,
Process_name,  received_date AS case_received_date,
ecm.created_date AS case_created_date,
ecff.CASE_FLEX_FIELD1,ecff.CASE_FLEX_FIELD2,ecff.CASE_FLEX_FIELD3,ecff.CASE_FLEX_FIELD4,
ecff.CASE_FLEX_FIELD5,ecff.CASE_FLEX_FIELD6,ecff.CASE_FLEX_FIELD7,ecff.CASE_FLEX_FIELD8,
ecff.CASE_FLEX_FIELD9,ecff.CASE_FLEX_FIELD10,ecff.CASE_FLEX_FIELD11,ecff.CASE_FLEX_FIELD12,
ecff.CASE_FLEX_FIELD13,ecff.CASE_FLEX_FIELD14,ecff.CASE_FLEX_FIELD15,ecff.CASE_FLEX_FIELD16,
ecff.CASE_FLEX_FIELD17,ecff.CASE_FLEX_FIELD18,ecff.CASE_FLEX_FIELD19,ecff.CASE_FLEX_FIELD20,
ecff.CASE_FLEX_FIELD21,ecff.CASE_FLEX_FIELD22,ecff.CASE_FLEX_FIELD23,ecff.CASE_FLEX_FIELD24,
ecff.CASE_FLEX_FIELD25,ecff.CASE_FLEX_FIELD26,ecff.CASE_FLEX_FIELD27,ecff.CASE_FLEX_FIELD28,
ecff.CASE_FLEX_FIELD29,ecff.CASE_FLEX_FIELD30,
checkpoint_live.fn_case_status(case_status) AS 'Case_status',
checkpoint_live.fn_check_status(check_status) AS 'check_status',
ec.check_name AS 'Check_unique_name',
check_disposition_id,disposition_name,check_severity,
REPLACE(ecc.closure_comments,'rn','') AS closure_comments,ecc.check_closure_date,
ecc.insuff_remarks,

@insuffdate:=(SELECT action_taken_on FROM ec_case_history ech WHERE ecc.case_check_id=ech.check_id
AND action_taken in (' case Status changed to : Case Insufficient',
' case Status changed to : InSufficient',
'Insuff Raised',
'Marked Insufficient',
'Marked Insufficient (Parallel Research)',
'New Status - Case Insufficient',
'New Status - InSufficient',
'Vendor request closed & Insuff raised',
'Case Insufficient',
'Check Created | Marked Insufficient',
'Check Updated | Marked Insufficient',
'New Status - Case Insuff Updated',
'New Status - New Case | Case Insufficient',
'Check Insuff raised',
'Case level Insuff raised',
'New Status - Case Insufficiency raised',
'Insufficient - Rework on Report',
'Insuff accepted' ) ORDER BY action_id LIMIT 1) AS 'First_Insuff_Date',

ecc.insuff_fulfill_date,
office_name location,
IF(ecc.family_id=4,emei.institute_name,IF(ecc.family_id=5,emc1.company_name,emcy.city_name)) AS verification_source,
ecm.case_expected_closure_date case_due_date,check_created_on,ecc.go_ahead_date,ecc.copy_of_check,

IFNULL(ec.CHECK_OPS_NAME,LEFT(REPLACE(family_name,' Family',''),3)) AS 'Check Ops Name',

ecc.reopen_date AS 'check_reopen_date',ecm.reopen_date AS 'case_reopen_date',
checkpoint_live.fn_ver_summary(ecc.VER_SUMMARY) AS Ver_Summary,
checkpoint_live.fn_ver_procedure(ecc.VERIFICATION_PROCEDURE) AS VERIFICATION_PROCEDURE,
checkpoint_live.fn_check_sub_status(ecc.sub_status) AS 'Check Sub Status',
(SELECT action_taken_on FROM ec_case_history ech
WHERE ecc.case_check_id=ech.check_id AND action_taken='Insuff Qc Status Updateas'
AND action_comments='Accepted' ORDER BY action_id DESC LIMIT 1) AS 'Inusff Accepted',

(SELECT report_sent_on FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=0 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Inerim_Report_Sent_Date',
(SELECT report_severity FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=0 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Inerim_Report_Severity',
(SELECT report_sent_on FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=1 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Final_Report_Sent_Date',
(SELECT report_severity FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=1 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Inerim_Report_Severity',
dqc_released_date,
(CASE TIER
WHEN 0 THEN 'Overseas'
WHEN 1 THEN 'Tier 1'
WHEN 2 THEN 'Tier 2'
WHEN 3 THEN 'Tier 3'
WHEN 4 THEN 'Tier 4' END ) AS 'Tier',

CONCAT(eud1.user_first_name,' ',eud1.user_last_name) AS 'DS Name',
QUEUE_NAME,
IF(@insuffdate is not null,if(@insuffdate<=dqc_released_date,'L1','L2'),'') AS 'Insuff Type',
IF(@insuffdate is not null,if(check_status in (0,1,2,3,4,5,6,7,13),'WIP','Non-Wip'),'Others') AS 'Insuff WIP Type',
PRIORITIZED_REQUESTED_EDC as 'EDC Prioritized Requested Date',
PRIORITIZED_REVISED_EDC as 'EDC Prioritized Revised Date',
if(ecc.check_status in (8,10,11,12),fn_user_name(VQC_REVIEWER_ID),'') as 'VQC done by',
ecm.CLIENT_CASE_EXPECTED_CLOSURE_DATE

FROM ec_case_master ecm
LEFT JOIN ec_case_fields ecff ON ecm.case_id=ecff.case_id
LEFT JOIN ec_case_checks ecc ON ecm.case_id = ecc.case_id
LEFT JOIN ec_check_queues ecq ON ecc.check_queue=ecq.queue_id AND ecc.check_id=ecq.check_id
LEFT JOIN ec_master_company emc ON emc.company_id = ecm.client_id
LEFT JOIN ec_client_process ecp ON ecm.process_id=ecp.process_id
LEFT JOIN ec_case_candidates ecc1 ON ecc1.candidate_id=ecm.candidate_id
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id=emcl.office_id
LEFT JOIN ec_user_details eud ON ecc.check_verifier=eud.user_id
LEFT JOIN ec_user_details eud1 ON ecm.documented_by=eud1.user_id
LEFT JOIN ec_case_check_verification_source eccvs ON ecc.case_check_id=eccvs.case_check_id
LEFT JOIN ec_master_company emc1 ON eccvs.org_id=emc1.company_id
LEFT JOIN ec_master_educational_institute emei ON eccvs.org_id=emei.institute_id
LEFT JOIN ec_master_city emcy ON emcy.city_id=eccvs.org_id
LEFT JOIN ec_master_state ems ON emcy.state_id=ems.state_id
LEFT JOIN ec_checks ec ON ecc.check_id=ec.check_id
left join ec_check_families ecf on ec.family_id=ecf.family_id
LEFT JOIN ec_master_disposition emd ON ecc.check_disposition_id=emd.disposition_id
WHERE case_status NOT IN (8,14)
AND check_status <> 9
and ecc.check_id<>193
AND ecm.case_ars_no IN  (check_id1)
"""

CASE_HISTORY_QUERY = """
SELECT company_name CLIENT,case_ars_no,check_name,(CASE check_status
WHEN '0' THEN 'Documentation Pending'
WHEN '1' THEN 'New/UnAssigned'
WHEN '2' THEN 'On Hold'
WHEN '3' THEN 'Insufficient'
WHEN '4' THEN 'Work in Progress'
WHEN '5' THEN 'Awaiting Response'
WHEN '6' THEN 'Escalated'
WHEN '7' THEN 'In Research'
WHEN '8' THEN 'Completed'
WHEN '9' THEN 'Disabled'
WHEN '10' THEN 'case closed by client'
WHEN '11' THEN 'Closed with Insufficiency'
WHEN '12' THEN 'Closed-Case Insufficient'
WHEN '13' THEN 'Contractually on Hold'END) AS 'check_status',check_severity,CONCAT(user_first_name,' ',user_last_name) AS action_taken_by, IF(eud.reporting_to IN (70,34,79,650),'insuff raised by PVT','insuff raised by other department') AS insuff_raised_by,designation,ech.*
FROM ec_case_history ech
LEFT JOIN ec_case_checks ecc ON  ech.check_id=ecc.case_check_id
LEFT JOIN ec_case_master ecm ON ech.case_id=ecm.case_id
LEFT JOIN ec_master_company emc ON ecm.client_id=emc.company_id
LEFT JOIN ec_user_details eud ON ech.action_taken_by=eud.user_id
WHERE case_ars_no in (check_id1)
"""

SENT_CASES_QUERY = """
SELECT ecr.case_report_id,company_name AS Client_name,
CONCAT(ifnull(first_name,''),' ',ifnull(Middle_name,''),' ',ifnull(Last_name,'')) Candidate_name,
Process_name,office_name location,case_ars_no,received_date,ecm.created_date,requested_on,ecm.case_expected_closure_date,
ecf3.case_flex_field1,ecf3.case_flex_field2,ecf3.case_flex_field3,
ecf3.case_flex_field4,ecf3.case_flex_field5,ecf3.case_flex_field6,
report_delivery_date, report_sent_on,
(CASE report_type
WHEN '0' THEN 'Interim'
WHEN '1' THEN 'Final'
WHEN '2' THEN 'Additional' END) AS 'Report Type',
report_severity, CONCAT(User_first_name,' ',User_last_name) PS,
(CASE report_status
WHEN 1 THEN 'New'
WHEN 2 THEN 'Assigned'
WHEN 3 THEN 'Verified'
WHEN 4 THEN 'Sent for Rework'
WHEN 5 THEN 'sent'
WHEN 6 THEN 'Reworked'
WHEN 7 THEN 'Not Sent'
WHEN 8 THEN 'Report Generation Pending'
WHEN 9 THEN 'Report Generation Failed'
WHEN 10 THEN 'Report Email Auto Triggered'
WHEN 11 THEN 'Report Email Sending Process Failed' END) AS 'Report Status',
CASE_ACCEPTED_DATE,ecr.requested_on report_triggered_on,spoc_comments cat_comments,
REVIEW_COMMENTS as 'ps comments',REVIEW_DATETIME as 'ps comment on',
SPOC_COMMENTS as cat_comments,CLIENT_CASE_EXPECTED_CLOSURE_DATE
FROM ec_case_reports ecr
LEFT JOIN ec_case_master ecm ON  ecr.case_id=ecm.case_id
LEFT JOIN ec_case_candidates ec ON ec.candidate_id=ecm.candidate_id
LEFT JOIN ec_master_company emc ON ecr.CLIENT_ID = emc.COMPANY_ID
LEFT JOIN ec_user_details eud ON ecr.assigned_to=eud.user_id
LEFT JOIN ec_client_process ecp ON ecm.process_id=ecp.process_id
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id=emcl.office_id
LEFT JOIN ec_case_fields ecf3 ON ecm.case_id=ecf3.case_id
WHERE  ecm.case_ars_no in (check_id1)
"""


def get_db_connection():
    """Open a fresh connection to the live checkpoint_live MySQL DB (RDS).
    Each of the 3 ARS queries below opens, executes, and closes its own
    connection rather than sharing one long-lived session — mirrors how the
    old Playwright flow made 3 independent requests in one browser session,
    without needing to keep a DB connection alive across the whole Streamlit
    button-click lifecycle."""
    return mysql.connector.connect(
        host=DB_HOST,
        port=DB_PORT,
        database=DB_NAME,
        user=DB_USER,
        password=DB_PASSWORD,
    )


def _execute_ars_query(sql_template, ars_numbers, extra_where=None, extra_params=()):
    """Run one of the *_QUERY templates above against checkpoint_live,
    substituting `check_id1` for a parameterized `%s,%s,...` IN-list bound to
    ars_numbers (never string-interpolated). Returns (header, data_rows, error)
    — header is the list of column names from cursor.description, data_rows a
    list of tuples — the same (header, rows) shape the old CSV-export parsing
    used to hand back, so every downstream add_*_sheet function below needs
    only to accept rows directly instead of a file path, not a logic rewrite.

    `extra_where`, if given, is appended as `AND (<extra_where>)` onto the
    query's existing WHERE clause (each *_QUERY template's WHERE is its last
    clause, with nothing after it, so straight string concatenation is safe)
    — used to push a filter down to the DB instead of fetching every row and
    filtering in Python. Deliberately NOT done by wrapping the query in a
    `SELECT * FROM (...) t WHERE ...` derived table: CASE_HISTORY_QUERY's own
    SELECT list has two columns that collide case-insensitively (`action_taken_by`
    or the explicit alias, `ACTION_TAKEN_BY` from `ech.*`), which MySQL accepts
    in a flat result set but rejects when materializing a derived table
    (`1060: Duplicate column name 'ACTION_TAKEN_BY'`, confirmed live). Flat
    string concatenation onto the existing WHERE has no such restriction.
    `extra_params` are appended after `ars_numbers` in the bound parameter
    list, in the same order their `%s` placeholders appear in `extra_where`."""
    placeholders = ",".join(["%s"] * len(ars_numbers))
    query = sql_template.replace("check_id1", placeholders)
    if extra_where:
        query = f"{query.rstrip()}\nAND ({extra_where})"
    params = list(ars_numbers) + list(extra_params)

    try:
        conn = get_db_connection()
    except MySQLError as e:
        return None, None, f"Database connection error: {e}"

    try:
        cursor = conn.cursor()
        cursor.execute(query, params)
        header = [desc[0] for desc in cursor.description]
        data_rows = cursor.fetchall()
        cursor.close()
        return header, data_rows, None
    except MySQLError as e:
        return None, None, f"Database query error: {e}"
    finally:
        conn.close()


def _excel_safe(value):
    """Convert a DB-native value into something openpyxl can write directly.
    decimal.Decimal (returned for DECIMAL/NUMERIC columns) isn't one of
    openpyxl's accepted cell types and raises on assignment — everything else
    (str/int/float/date/datetime/None) is already safe as-is."""
    if isinstance(value, Decimal):
        return float(value)
    return value


def _none_first_sort_key(value):
    """Sort key for a column that may hold a real date/datetime or None —
    plain `sorted()` raises TypeError comparing None to a datetime, so this
    sorts blanks first instead of erroring."""
    return (value is None, value)


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


def add_red_remarks_sheet(wb, header, data_rows):
    """Filter the Advance Tracker query results to check_severity in
    RED_REMARKS_SEVERITIES AND Check_unique_name in RED_REMARKS_CHECK_NAMES,
    then group by case_ars_no and join each ARS's closure_comments with ", "
    into a single row per unique ARS. Writes case_ars_no (column B) ->
    combined closure_comments (column C) into a new 'red remarks' sheet in
    the cleaned tracker — verified against a real export to reproduce the
    expected combined-comment text exactly (e.g. ARS 3055-016865).

    Takes an already-open Workbook (`wb`) rather than opening/saving its own
    copy of the file — see the "single load/save pass" note on the button
    handler below for why: opening and saving this file separately per
    function meant 6 full load+parse+save round trips per run on a workbook
    that can hold tens of thousands of audit-sheet rows, which dominated this
    step's wall-clock time far more than the DB queries themselves."""
    required = ["check_severity", "Check_unique_name", "case_ars_no", "closure_comments"]
    missing = [c for c in required if c not in header]
    if missing:
        return None, f"Column(s) {', '.join(missing)} not found in the query result"

    sev_col = header.index("check_severity")
    name_col = header.index("Check_unique_name")
    ars_col = header.index("case_ars_no")
    comment_col = header.index("closure_comments")

    filtered_rows = [
        row
        for row in data_rows
        if row[sev_col] in RED_REMARKS_SEVERITIES and row[name_col] in RED_REMARKS_CHECK_NAMES
    ]

    comments_by_ars = {}
    for row in filtered_rows:
        if not row[comment_col]:
            continue
        comments_by_ars.setdefault(row[ars_col], []).append(row[comment_col])

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

    return len(comments_by_ars), None


# =========================================================
# FUNCTION TO FILTER/SORT CASE HISTORY DATA INTO A "Form submisison - L2" TAB
# =========================================================
FORM_SUBMISSION_L2_COMMENT = (
    "Check moved to WIP. Antecedents populated from iBridge candidate submission."
)


def add_form_submission_l2_sheet(wb, header, data_rows):
    """Filter the Case History query results to ACTION_COMMENTS matching the
    L2 form-submission comment, sort by ACTION_TAKEN_ON oldest-first, and copy
    the full matching rows into a 'Form submisison - L2' sheet in the cleaned
    tracker. Takes an already-open Workbook — see add_red_remarks_sheet's
    docstring for why."""
    if "ACTION_COMMENTS" not in header or "ACTION_TAKEN_ON" not in header:
        return None, "Column 'ACTION_COMMENTS' or 'ACTION_TAKEN_ON' not found in the query result"

    comments_col = header.index("ACTION_COMMENTS")
    taken_on_col = header.index("ACTION_TAKEN_ON")

    filtered_rows = [row for row in data_rows if row[comments_col] == FORM_SUBMISSION_L2_COMMENT]
    filtered_rows.sort(key=lambda row: _none_first_sort_key(row[taken_on_col]))

    sheet_name = "Form submisison - L2"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    ws.append(header)
    for row in filtered_rows:
        ws.append([_excel_safe(v) for v in row])

    tracker_ws = wb["Pre_Offer_checkwise"]
    tracker_ws["X2"] = "Form submisison - L2"
    for row in range(3, tracker_ws.max_row + 1):
        tracker_ws[f"X{row}"] = (
            f"=IFERROR(VLOOKUP(C{row},'Form submisison - L2'!B:T,12,0),\"\")"
        )

    return len(filtered_rows), None


# =========================================================
# FUNCTION TO FILTER/SORT CASE HISTORY DATA INTO A "L2 check addition (Oldest)" TAB
# =========================================================
CASE_REOPENED_ACTION = "New Status - Case Reopened"


def add_l2_check_addition_oldest_sheet(wb, header, data_rows):
    """Filter the Case History query results to ACTION_TAKEN == 'New Status - Case
    Reopened', sort by ACTION_TAKEN_ON oldest-first, and copy the full matching
    rows into a 'L2 check addition (Oldest)' sheet in the cleaned tracker.
    Takes an already-open Workbook — see add_red_remarks_sheet's docstring
    for why."""
    if "ACTION_TAKEN" not in header or "ACTION_TAKEN_ON" not in header:
        return None, "Column 'ACTION_TAKEN' or 'ACTION_TAKEN_ON' not found in the query result"

    action_col = header.index("ACTION_TAKEN")
    taken_on_col = header.index("ACTION_TAKEN_ON")

    filtered_rows = [row for row in data_rows if row[action_col] == CASE_REOPENED_ACTION]
    filtered_rows.sort(key=lambda row: _none_first_sort_key(row[taken_on_col]))

    sheet_name = "L2 check addition (Oldest)"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    ws.append(header)
    for row in filtered_rows:
        ws.append([_excel_safe(v) for v in row])

    tracker_ws = wb["Pre_Offer_checkwise"]
    tracker_ws["Y2"] = "L2 check addition (Oldest)"
    for row in range(3, tracker_ws.max_row + 1):
        tracker_ws[f"Y{row}"] = (
            f"=IFERROR(VLOOKUP(C{row},'L2 check addition (Oldest)'!B:T,12,0),\"\")"
        )

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


def add_l2_report_sent_sheet(wb, header, data_rows):
    """Filter the Sent Cases query results to Report Status == 'Sent' and
    Report Type == 'Additional', sort by report_sent_on oldest-first, and copy
    the full matching rows into a 'L2 Report Sent' sheet in the cleaned tracker.
    Overwrites the existing 'L2 Report sent date' / 'L2 Report sent severity'
    values in Pre_Offer_checkwise (every row 3..max_row gets a fresh formula
    below), mapping case_ars_no -> ARS Number (column C) to refill them
    (date-only, dd-mmm-yy, for the date column).

    Uses INDEX/MATCH rather than VLOOKUP for the mapping formulas: unlike the
    other query-derived sheets here, the relative position of case_ars_no vs.
    report_sent_on/report_severity in the Sent Cases result hasn't been verified
    against a live file, and VLOOKUP can't look left of its lookup column.
    Takes an already-open Workbook — see add_red_remarks_sheet's docstring
    for why."""
    required = ["Report Status", "Report Type", "report_sent_on", "case_ars_no", "report_severity"]
    missing = [c for c in required if c not in header]
    if missing:
        return None, f"Column(s) {', '.join(missing)} not found in the query result"

    status_col = header.index("Report Status")
    type_col = header.index("Report Type")
    sent_on_col = header.index("report_sent_on")
    ars_col = header.index("case_ars_no")
    severity_col = header.index("report_severity")

    filtered_rows = [
        row
        for row in data_rows
        if (row[status_col] or "").strip().lower() == REPORT_SENT_STATUS.lower()
        and (row[type_col] or "").strip().lower() == REPORT_SENT_TYPE.lower()
    ]
    filtered_rows.sort(key=lambda row: _none_first_sort_key(row[sent_on_col]))

    sheet_name = "L2 Report Sent"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    ws.append(header)
    for row in filtered_rows:
        ws.append([_excel_safe(v) for v in row])

    # Reduce report_sent_on to a real date (date-only) in the audit sheet so
    # the INDEX/MATCH pull-through below carries an actual date value, not a
    # datetime with a time component — the destination's dd-mmm-yy number
    # format needs a date-typed value, and the query returns a full datetime.
    sent_on_letter = get_column_letter(sent_on_col + 1)
    for i, row in enumerate(filtered_rows, start=2):
        raw_value = row[sent_on_col]
        parsed_date = None
        if isinstance(raw_value, datetime):
            parsed_date = raw_value.date()
        elif isinstance(raw_value, date):
            parsed_date = raw_value
        elif isinstance(raw_value, str) and raw_value.strip():
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

    return len(filtered_rows), None


# =========================================================
# FUNCTION TO UPDATE "Case Status" WHEN "L2 Report sent date" IS POPULATED
# =========================================================
CASE_STATUS_TARGETS = {"Insufficient", "On Hold"}
CASE_STATUS_COMPLETED = "Completed"


def update_case_status_from_l2_report_sent(wb):
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
    if the L2 Report Sent data changes) never touch it again. Takes an
    already-open Workbook — see add_red_remarks_sheet's docstring for why."""
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

    return updated_count, None


# =========================================================
# FUNCTION TO ADD L2 DUE DATE / L2 TAT / L2 CHECK STATUS COLUMNS
# =========================================================
def add_l2_summary_columns(wb):
    """Add 'L2 Due Date' (Z), 'L2 TAT' (AA), and 'L2 check status' (AB) columns to
    Pre_Offer_checkwise. Z is derived from X ('Form submisison - L2', added by
    add_form_submission_l2_sheet) via WORKDAY.INTL against the 'a' named range
    (the holiday calendar written by _write_holiday_list_sheet), and AA from
    Z and U ('L2 Report sent date'). AB checks X, Y ('L2 check addition
    (Oldest)', added by add_l2_check_addition_oldest_sheet — an independent
    lookup from X, not derived from it, unlike Z), U, and K ('Case Status').
    Must run after add_form_submission_l2_sheet (X), add_l2_check_addition_oldest_sheet
    (Y), and clean_latest_excel_file (the 'a' named range). Takes an
    already-open Workbook — see add_red_remarks_sheet's docstring for why."""
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
st.subheader("Step 3: Run ARS Query")
st.caption(
    "Queries the checkpoint_live database directly for Advance Tracker, Case "
    "History, and Sent Cases data for the ARS numbers — no browser involved"
)

if st.button("🔎 Run ARS Query", type="primary", use_container_width=True):
    ars_numbers, error = get_ars_numbers()
    cleaned_file_path = DOWNLOAD_DIR / CLEANED_FILE_NAME

    if error:
        st.error(error)
    else:
        try:
            with st.spinner("Querying Advance Tracker..."):
                adv_header, adv_rows, adv_error = _execute_ars_query(
                    ADVANCE_TRACKER_QUERY, ars_numbers
                )
            if adv_error:
                st.error(adv_error)
            else:
                st.success(f"Fetched {len(adv_rows)} Advance Tracker row(s) for {len(ars_numbers)} ARS number(s)")

            # Case History is pushed a server-side filter for the only two
            # conditions any downstream sheet actually uses (ACTION_COMMENTS
            # == the L2 form-submission comment, or ACTION_TAKEN == the case
            # reopened action) instead of fetching every row and filtering in
            # Python. Verified against the live DB to return byte-identical
            # rows either way, while cutting a 591K-row/16s fetch down to
            # ~16K rows/<1s at production scale — see CLAUDE.md.
            with st.spinner("Querying Case History..."):
                ch_header, ch_rows, ch_error = _execute_ars_query(
                    CASE_HISTORY_QUERY,
                    ars_numbers,
                    extra_where="ech.ACTION_COMMENTS = %s OR ech.ACTION_TAKEN = %s",
                    extra_params=[FORM_SUBMISSION_L2_COMMENT, CASE_REOPENED_ACTION],
                )
            if ch_error:
                st.error(ch_error)
            else:
                st.success(f"Fetched {len(ch_rows)} relevant Case History row(s)")

            with st.spinner("Querying Sent Cases..."):
                sc_header, sc_rows, sc_error = _execute_ars_query(SENT_CASES_QUERY, ars_numbers)
            if sc_error:
                st.error(sc_error)
            else:
                st.success(f"Fetched {len(sc_rows)} Sent Cases row(s)")

            # One load/mutate/save pass instead of a separate open+save per
            # step — opening and saving this workbook 6 separate times (once
            # per add_*_sheet/update/summary call) dominated this button's
            # wall-clock time far more than the DB queries themselves once
            # Case History is filtered server-side. See each function's
            # docstring for why it now takes `wb` instead of opening its own.
            with st.spinner("Writing results into the tracker..."):
                wb = load_workbook(cleaned_file_path)

                if not adv_error:
                    red_remarks_count, red_remarks_error = add_red_remarks_sheet(
                        wb, adv_header, adv_rows
                    )
                    if red_remarks_error:
                        st.error(red_remarks_error)
                    else:
                        st.success(
                            f"Added 'red remarks' tab with {red_remarks_count} Major Discrepancy row(s)"
                        )

                if not ch_error:
                    l2_count, l2_error = add_form_submission_l2_sheet(wb, ch_header, ch_rows)
                    if l2_error:
                        st.error(l2_error)
                    else:
                        st.success(f"Added 'Form submisison - L2' tab with {l2_count} row(s)")

                    reopened_count, reopened_error = add_l2_check_addition_oldest_sheet(
                        wb, ch_header, ch_rows
                    )
                    if reopened_error:
                        st.error(reopened_error)
                    else:
                        st.success(
                            f"Added 'L2 check addition (Oldest)' tab with {reopened_count} row(s)"
                        )

                if not sc_error:
                    sent_count, sent_error = add_l2_report_sent_sheet(wb, sc_header, sc_rows)
                    if sent_error:
                        st.error(sent_error)
                    else:
                        st.success(
                            f"Added 'L2 Report Sent' tab with {sent_count} row(s) and mapped "
                            "'L2 Report sent date' / 'L2 Report sent severity'"
                        )

                status_count, status_error = update_case_status_from_l2_report_sent(wb)
                if status_error:
                    st.error(status_error)
                else:
                    st.success(f"Updated Case Status to 'Completed' for {status_count} row(s)")

                summary_success, summary_error = add_l2_summary_columns(wb)
                if summary_success:
                    st.success("Added 'L2 Due Date', 'L2 TAT', and 'L2 check status' columns")
                else:
                    st.error(summary_error)

                wb.save(cleaned_file_path)

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
            st.error(f"Error running ARS query: {str(e) or type(e).__name__}")
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
