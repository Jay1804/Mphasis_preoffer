# Mphasis Tracker

Single-file Streamlit app (`mphasis.py`) that downloads the latest Mphasis ARS
tracker Excel file from Outlook, cleans and annotates it, then queries the
`checkpoint_live` MySQL database directly to pull supplementary Advance
Tracker / Case History / Sent Cases data and merge it back in.

## Run

```
streamlit run mphasis.py
```

Requires Outlook to be running on the same Windows machine (uses `win32com.client`
COM automation, not the Graph API). Also requires network access to the
`checkpoint_live` MySQL DB (an AWS RDS instance) and the `mysql-connector-python`
package — see "MIS data retrieval: direct DB query" below. Also requires a real
installed copy of Excel on the same machine — `add_status_view_sheet()` drives
it via COM automation (a hidden instance) to build genuine, live PivotTables,
since openpyxl cannot create real PivotTable objects at all.

Historical note: this step originally drove the AuthBridge MIS export-query
website (`https://mis.authbridge.com/export_query/`) via browser automation
— first Selenium, later Playwright — before being replaced with a direct DB
connection (see below). Neither Playwright nor a Chrome browser is required
anymore for this step.

## Flow

1. **Manual trigger** — user opens `TRIGGER_URL` in a browser to kick off the
   tracker email generation (external system, not controlled by this app).
2. **Download Latest** — `download_latest_from_email()` polls an Outlook Inbox
   for the newest email whose subject contains the configured keyword, and saves
   all its attachments into `DOWNLOAD_DIR`. Which Inbox is picked is controlled
   by the "Outlook account" UI field via `get_target_inbox()` — see gotcha below.
3. **Clean Latest File** — `clean_latest_excel_file()` takes the most-recently
   modified `.xlsx`/`.xls` source in `DOWNLOAD_DIR` (via `get_latest_excel_file()`,
   which excludes its own output) and:
   - deletes specific columns by header name (see below)
   - inserts "Due Date" (N) and "L1 TAT" (O) columns with formulas + `dd-mmm-yy`
     date formatting, filled from row 3 to the last row. N's formula is
     `=IF(G{row}<>"",WORKDAY.INTL(G{row},1,11,a),WORKDAY.INTL(E{row},1,11,a))`
     — both branches weigh against the `a` named range (the holiday calendar
     written by `_write_holiday_list_sheet`, called later in this same
     function but before `wb.save()` — order-independent since `a` just needs
     to exist in the saved file, not before this formula text is written).
     Updated at explicit user request to add the `,a` holiday argument to
     both branches — it previously used the bare 3-arg `WORKDAY.INTL(...,11)`
     form with no holiday list at all, unlike Z ("L2 Due Date", see below)
     which always included it. Don't revert this back to the 3-arg form.
     Verified live via Excel COM (`CalculateFullRebuild()`, same technique
     `add_status_view_sheet` uses): formula text matches exactly, `a`
     resolves to all 216 holiday dates, and sample rows evaluated to correct
     real dates (e.g. `G=10-Aug-2026` → `N=11-Aug-2026`, the next workday).
   - fills column **AC** ("Ars N01" header) with a quoted/comma-joined ARS
     number list formula (`="'"&C{row}&"'"&IF(C{row+1}<>"",",","")`) — each row
     checks the *next* row's column C to decide the trailing comma
   - saves the result as `Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx` in the same directory
4. **Run ARS Query** — one button handler chains: `_execute_ars_query(ADVANCE_TRACKER_QUERY, ...)`
   → `add_red_remarks_sheet` → `_execute_ars_query(CASE_HISTORY_QUERY, ...)` →
   `add_form_submission_l2_sheet` → `add_l2_check_addition_oldest_sheet` (both off
   the same Case History result — one query, reused for both sheets) →
   `_execute_ars_query(SENT_CASES_QUERY, ...)` → `add_l2_report_sent_sheet` →
   `update_case_status_from_l2_report_sent` → `add_l2_summary_columns` →
   `add_status_view_sheet` → `apply_professional_styling`.
   See "MIS data retrieval: direct DB query" below. (`clean_latest_excel_file`, step 3, also calls `_write_holiday_list_sheet()`
   to set up the `Holiday List` sheet + `a` named range that both its own N
   ("Due Date") formula and `add_l2_summary_columns`'s Z formula depend on —
   see the Holiday List gotcha below. It also calls
   `add_status_view_sheet()` itself, right after cleaning, so the report exists —
   with its L2 blocks empty — even before any MIS query has run.)
5. **Download Final Report** — a `st.download_button` reads `Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx`
   straight off disk (`final_report_path.read_bytes()`) so the user can grab it
   without opening `DOWNLOAD_DIR` in Explorer. Re-reads the file fresh on every
   Streamlit rerun, so it always reflects the latest save, and hides itself if
   the file doesn't exist yet.

`apply_professional_styling()` is a purely cosmetic finalization pass —
bold white-on-navy header rows, frozen header panes, readable column widths,
hidden gridlines + a distinguishing tab color on the audit sheets, text
wrapping for `red remarks`!`closure_comments`, and hidden gridlines + a
frozen pane (`A5`) on `Status view` — run twice: once right after
**Clean Latest File** (so `Pre_Offer_checkwise` looks polished even before
any MIS query runs) and again as the last step of **Run ARS Query in MIS**
(so it also covers the sheets/headers that step creates). It never touches
cell values or formulas, only styling/dimensions — see `_style_header_row()`.
Verified against the real `Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx`: header row 2 styled
uniformly across both original tracker columns and every column this app
adds, existing wider column widths preserved (the width logic only
increases, never shrinks), `red remarks` wraps correctly at an 80-wide
column C.

## Step 4 in detail: MIS data retrieval — direct DB query

This step used to drive the AuthBridge MIS export-query website via browser
automation (first Selenium, later Playwright) — logging in, selecting Host /
Database / Data Time Slab / Query dropdowns, clicking Export, and downloading
a CSV/zip. It now connects **directly to the `checkpoint_live` MySQL DB**
(an AWS RDS instance) and runs the same three SQL queries the website ran
server-side, via `mysql-connector-python`. Same data, same column names, no
browser, no MIS website dependency at all. The Playwright/Selenium era is
summarized as history near the end of "Key gotchas" below, in case browser
automation ever needs reviving; none of that code exists in `mphasis.py`
anymore.

**Credentials & connection** (`DB_HOST`, `DB_PORT`, `DB_NAME`, `DB_USER`,
`DB_PASSWORD`, all from `.env` — see the credentials gotcha below):
`get_db_connection()` opens a fresh `mysql.connector` connection per query
(not pooled/shared across the three queries — mirrors how the old Playwright
flow made 3 independent requests in one browser session, without needing to
keep a DB connection alive across the whole Streamlit button-click
lifecycle).

**The three queries** (`ADVANCE_TRACKER_QUERY`, `CASE_HISTORY_QUERY`,
`SENT_CASES_QUERY` in `mphasis.py`) are stored **verbatim** as given — same
joins, same `CASE WHEN` status/severity mappings, same subqueries for
`First_Insuff_Date`/`Last_Inerim_Report_Sent_Date`/etc. `check_id1` is the
literal placeholder text from the site's own SQL (its ARS-number-list
substitution point) — kept as-is rather than rewritten to a named parameter,
and swapped for a parameterized `%s,%s,...` IN-list at execute time by
`_execute_ars_query(sql_template, ars_numbers)`, which does
`sql_template.replace("check_id1", placeholders)` then
`cursor.execute(query, ars_numbers)` — ARS numbers are always bound as query
parameters, never string-interpolated into the SQL text. Returns
`(header, data_rows, error)`: `header` is `[d[0] for d in cursor.description]`,
`data_rows` a list of tuples — the same `(header, rows)` shape
`_read_export_csv_rows()` used to hand back from a parsed CSV, so every
`add_*_sheet` function below only needed to accept rows directly instead of a
file path, not a logic rewrite. **Verified against the live DB** at full
production scale (3,395 ARS numbers): Advance Tracker returned 21,001 rows,
Case History 591,328 rows (see the performance note directly below — Case
History is no longer fetched unfiltered in the actual app), Sent Cases 6,964
rows — all three queries, and every downstream
`add_*_sheet`/`update_case_status_from_l2_report_sent`/`add_l2_summary_columns`
step, ran end-to-end successfully against real production data before this
rewrite was considered done.

**Performance: Case History is filtered server-side, not in Python.**
`_execute_ars_query()` takes an optional `extra_where`/`extra_params` pair,
appended as `AND (<extra_where>)` onto the query template's existing WHERE
clause (each `*_QUERY` template's WHERE is its last clause with nothing
after it, so plain string concatenation is safe — no derived-table wrapping
needed, see below for why that specifically doesn't work here). The button
handler calls `CASE_HISTORY_QUERY` with
`extra_where="ech.ACTION_COMMENTS = %s OR ech.ACTION_TAKEN = %s"` and
`extra_params=[FORM_SUBMISSION_L2_COMMENT, CASE_REOPENED_ACTION]` — the exact
two conditions `add_form_submission_l2_sheet`/`add_l2_check_addition_oldest_sheet`
already filtered for in Python, now pushed down to MySQL instead. **Verified
both for correctness and speed** against the live DB (497 ARS numbers): the
SQL-filtered result and the old fetch-everything-then-filter-in-Python result
are byte-identical row sets (0 rows differing either way), while the fetch
itself dropped from 16.2s/108,508 rows to 0.92s/2,313 rows. At full
production scale (3,395 ARS numbers) this took Case History from 591,328
unfiltered rows down to 15,695 — an exact match for
`Form submisison - L2` (8,577) + `L2 check addition (Oldest)` (7,118) counts
confirmed in an earlier full run — while cutting that query's fetch time from
tens of seconds to ~6s. Don't revert this to fetching Case History
unfiltered "to keep the query simple" — the two-condition filter is exactly
what's consumed downstream, nothing is lost.

**Why this filter is appended directly to the flat query rather than via a
`SELECT * FROM (CASE_HISTORY_QUERY) t WHERE ...` wrapper** (which would have
kept `CASE_HISTORY_QUERY`'s own text even more clearly untouched): confirmed
live that the wrapper form fails with
`mysql.connector.errors.ProgrammingError: 1060 (42S21): Duplicate column
name 'ACTION_TAKEN_BY'`. `CASE_HISTORY_QUERY`'s `SELECT` list has both an
explicit `CONCAT(user_first_name,' ',user_last_name) AS action_taken_by` and
`ech.*` (which separately includes the raw `ACTION_TAKEN_BY` foreign-key
column) — MySQL column names are case-insensitive, so these two collide.
A flat `SELECT`'s result set tolerates duplicate/colliding column names
just fine (confirmed: `cursor.description` lists both `'action_taken_by'`
and `'ACTION_TAKEN_BY'` as separate entries), but materializing that same
result set as a derived table does not. Appending straight onto the
existing WHERE clause never creates a derived table, so this doesn't apply.
If a future query ever needs the derived-table form for some other reason,
expect this same collision.

**Two type-handling differences from the old CSV-parsing path**, both
confirmed live:
- **DB rows carry real Python types, not strings.** `ACTION_TAKEN_ON` and
  `report_sent_on` come back as real `datetime.datetime` objects (confirmed:
  `type(row[aton_col])` is `datetime.datetime`), not the zero-padded
  `YYYY-MM-DD HH:MM:SS` strings the old CSV export gave. Sorting by these
  columns now compares real datetimes directly (more robust than the old
  lexical string sort) — but plain `sorted()`/`.sort()` raises `TypeError`
  comparing `None` to a `datetime` if any row has a blank timestamp, so every
  sort here uses the `_none_first_sort_key()` helper (`(value is None,
  value)`) instead of a bare `row[col]` key.
- **`_excel_safe()` converts `decimal.Decimal` to `float`** before any DB row
  is written into an openpyxl sheet via `ws.append(...)` — DECIMAL/NUMERIC
  columns come back as `Decimal`, which openpyxl cannot write directly and
  raises on. `add_form_submission_l2_sheet`, `add_l2_check_addition_oldest_sheet`,
  and `add_l2_report_sent_sheet` all map every row through this before
  appending. `add_l2_report_sent_sheet`'s `report_sent_on`-to-date reduction
  (needed so the destination `dd-mmm-yy` number format applies) now branches
  on `isinstance(raw_value, datetime)` / `isinstance(raw_value, date)` first,
  falling back to the old string-parsing only if the value somehow arrives as
  a string.

The button handler:
1. Builds the raw ARS number list **directly from `Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx` column C**
   via `get_ars_numbers()` — not by reading the AC column's formula (openpyxl
   never evaluates formulas, so AC would just return unevaluated formula text).
   Returns a plain Python list (no manual quoting needed — `_execute_ars_query`
   binds it as parameters).
2. `_execute_ars_query(ADVANCE_TRACKER_QUERY, ars_numbers)`.
3. `_execute_ars_query(CASE_HISTORY_QUERY, ars_numbers, extra_where=..., extra_params=...)`
   — run **once** (server-side-filtered, see the performance note above), its
   `(header, rows)` reused for both `add_form_submission_l2_sheet` and
   `add_l2_check_addition_oldest_sheet` (avoids querying Case History twice
   for what used to be two separate exports of the same underlying query).
4. `_execute_ars_query(SENT_CASES_QUERY, ars_numbers)`.
5. **One `load_workbook()` → mutate → one `wb.save()`** for everything from
   here through `add_l2_summary_columns()`: `add_red_remarks_sheet`,
   `add_form_submission_l2_sheet`, `add_l2_check_addition_oldest_sheet`,
   `add_l2_report_sent_sheet`, `update_case_status_from_l2_report_sent`, and
   `add_l2_summary_columns` all now take an already-open `wb` as their first
   argument and neither open nor save the file themselves — the button
   handler does `wb = load_workbook(cleaned_file_path)` once, calls all six
   in sequence, then `wb.save(cleaned_file_path)` once. This replaced 6
   separate full load+parse+save round trips on a workbook that can hold
   tens of thousands of audit-sheet rows, which dominated this step's
   wall-clock time even more than the (now largely fixed) Case History
   over-fetch — confirmed live: the consolidated load+mutate+save pass alone
   still took ~57s at full production scale (3,395 ARS numbers, ~2,422
   tracker rows plus ~34K new audit-sheet rows across 4 sheets) — openpyxl's
   own XML serialization cost for a workbook this size, now paid once instead
   of 6 times. Each of the six functions' docstrings notes it takes `wb`
   instead of opening its own — don't revert any single one back to
   `load_workbook(cleaned_file_path)` / `wb.save(...)` internally, that
   reintroduces the redundant round trips for that function specifically.
6. Then `add_status_view_sheet()` → `apply_professional_styling()`,
   unchanged from before — these still do their own COM/openpyxl
   open+save (COM for the pivot build, openpyxl for the final styling pass),
   since each must run strictly after the previous step's save completes.

`add_red_remarks_sheet(wb, header, data_rows)`:
   - filters rows to `check_severity` in `RED_REMARKS_SEVERITIES` (`Major
     Discrepancy`, `Amber`, `Minor Discrepancy`) **AND** `Check_unique_name`
     in `RED_REMARKS_CHECK_NAMES` (`UAN Check for Undisclosed Employment`,
     `Dual Employment Verification via Form 26AS`, `Criminal Records
     Verification`, `National Identity Check`, `India Court Record Check
     through Law Firm`) — both exact match, both required. Confirmed live:
     `check_severity` values seen are `Green`, `None`, `Amber`, `Major
     Discrepancy`, `Minor Discrepancy` — matches the filter set exactly.
   - groups the filtered rows by `case_ars_no` and joins each ARS's
     `closure_comments` with `", "` into **one row per unique ARS** (an ARS
     can have multiple matching checks, e.g. both a Dual Employment and a UAN
     check) — verified to reproduce the exact expected combined-comment text
     against a real export for 16 sample ARS numbers (e.g. ARS
     `3055-016865` → `"Possible dual employment is found...., Verification
     could not be possible as the EPFO portal is not working...."`, joined in
     row-encounter order, not sorted)
   - writes `case_ars_no` → combined `closure_comments` into a new `red
     remarks` sheet in `Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx` (column B = `case_ars_no`,
     column C = combined `closure_comments`; replacing the sheet if it
     already exists) — a 3-column (A blank, B, C) aggregation, not a raw data
     dump of every matching row
   - fills `Pre_Offer_checkwise!W3:W{last row}` — the tracker's existing
     `Red Remarks` header column — with
     `=IFERROR(VLOOKUP(C{row},'red remarks'!B:AB,2,0),"")`, which looks up the
     ARS number against `red remarks` column B (`case_ars_no`) and returns
     column C (the combined `closure_comments`, the 2nd column of the B:AB
     range — the `AB` upper bound is arbitrary headroom, nothing needs to
     actually exist past column C for this to work)

`add_form_submission_l2_sheet(wb, header, data_rows)`:
   - filters the Case History rows to `ACTION_COMMENTS == FORM_SUBMISSION_L2_COMMENT`
     (`"Check moved to WIP. Antecedents populated from iBridge candidate
     submission."`), sorts by `ACTION_TAKEN_ON` oldest-first (real `datetime`
     comparison via `_none_first_sort_key()` — see the type-handling note above)
   - writes the result into a new `Form submisison - L2` sheet (that exact
     spelling/typo, kept verbatim per explicit instruction — don't "fix" it)
   - fills `Pre_Offer_checkwise!X2` with that same sheet name as the header, and
     `X3:X{last row}` with
     `=IFERROR(VLOOKUP(C{row},'Form submisison - L2'!B:T,12,0),"")` (column B =
     `case_ars_no`, the 12th column of B:T = `ACTION_TAKEN_ON`)

`add_l2_check_addition_oldest_sheet(wb, header, data_rows)`:
   - filters the same Case History rows to `ACTION_TAKEN == CASE_REOPENED_ACTION`
     (`"New Status - Case Reopened"`), same oldest-first sort by `ACTION_TAKEN_ON`
   - writes the result into a new `L2 check addition (Oldest)` sheet
   - fills `Pre_Offer_checkwise!Y2` with that same sheet name as the header, and
     `Y3:Y{last row}` with
     `=IFERROR(VLOOKUP(C{row},'L2 check addition (Oldest)'!B:T,12,0),"")` (same
     B:T/column-12 alignment as the `Form submisison - L2` lookup — column B =
     `case_ars_no`, the 12th column of B:T = `ACTION_TAKEN_ON`)

`add_l2_report_sent_sheet(wb, header, data_rows)`:
   - filters rows to `Report Status == "Sent"` AND `Report Type == "Additional"`
     (both exact match, case/whitespace-insensitive via `.strip().lower()`),
     sorts by `report_sent_on` oldest-first
   - writes the full matching rows into a new `L2 Report Sent` sheet
     (**column order matches the raw Sent Cases query's `SELECT` order** —
     `case_ars_no` lands at column **F**, index 5, not B; confirmed live —
     anything reading this sheet must look up columns by header name, never
     assume B), then reduces just that sheet's `report_sent_on` column to a
     real `date` value (needed so the destination `dd-mmm-yy` number format
     below actually applies — see the type-handling note above)
   - **clears** `Pre_Offer_checkwise`'s existing `L2 Report sent date` / `L2
     Report sent severity` column values first, then refills them via
     `INDEX/MATCH` (not `VLOOKUP`) against `L2 Report Sent`, matching on
     `case_ars_no` vs. `Pre_Offer_checkwise!C{row}`. INDEX/MATCH was chosen
     over VLOOKUP here specifically because `case_ars_no` is not known to sit
     to the left of `report_sent_on`/`report_severity` in that sheet (a hard
     requirement for VLOOKUP, but not for INDEX/MATCH) — confirmed true:
     `case_ars_no` (F) actually sits *right* of neither in this query's
     column order, so this wasn't just defensive. Both destination columns
     are located by **header name** via `_find_header_col_letter()` (scanning
     row 2) rather than hardcoded letters, for the same reason — see the
     debugging section below before assuming this is wired to the wrong
     columns. Confirmed live: `Report Status` values are `sent` (lowercase,
     matching the query's own `WHEN 5 THEN 'sent'` CASE branch), `Report
     Type` values seen are `Additional`/`Final` — filtering to
     `sent`+`additional` on a 6,964-row Sent Cases result yielded 3,551
     matching rows.

`update_case_status_from_l2_report_sent(wb)`:
    - for `Pre_Offer_checkwise` rows whose `Case Status` (K) is currently
      `Insufficient` or `On Hold` (both confirmed exact-match strings —
      the only other real values seen are `Completed` and `Work In
      Progress`), overwrites K to `Completed` if that row's ARS number
      appears in the `L2 Report Sent` sheet's `case_ars_no` column (looked
      up by header name — see `add_l2_report_sent_sheet`'s column-order
      gotcha above; a hardcoded `B` here was a real bug caught before it
      shipped), otherwise leaves K untouched
    - determines "L2 Report sent date is non-blank" from that same source
      data, **not** by reading `Pre_Offer_checkwise!U`'s cell value — U holds
      an `INDEX/MATCH` formula, and openpyxl never evaluates formulas, so
      reading it back would give `None`/unevaluated regardless of what it
      resolves to once opened in Excel
    - a one-time **value** overwrite, not a formula — `K{row}` holds original
      source data (not something this app derives), so "leave unchanged"
      only makes sense as a snapshot of whatever the cell currently holds; a
      formula can't reference its own cell's prior value
    - naturally idempotent across repeat runs: once a row flips to
      `Completed` it no longer matches `CASE_STATUS_TARGETS`, so a later run
      — even one where the `L2 Report Sent` data has changed — never touches
      it again (this is intentional per the literal spec: "else same as
      previous status" means leave it as whatever it currently is, not
      revert to some remembered original)
    - verified against the real `Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx`: 1770 of 2422 ARS
      numbers had an entry in `L2 Report Sent`; all 18 rows that were
      `Insufficient`/`On Hold` at the time happened to be among them, so all
      18 flipped to `Completed`; unrelated rows (already `Completed`/`Work
      In Progress`) were confirmed untouched
`add_l2_summary_columns(wb)` — runs last, since it depends on X
(`add_form_submission_l2_sheet`), Y (`add_l2_check_addition_oldest_sheet`),
U/"L2 Report sent date" (`add_l2_report_sent_sheet`), and the `a` named range
(written by `clean_latest_excel_file`, step 3) already being in place.
Z/AA depend on X only, not Y; **AB depends on both X and Y** — see the
AB bullet below for why that matters (it's not redundant with Z, despite
Y and Z both ultimately tracing back to a per-row lookup). Adds three
more `Pre_Offer_checkwise` columns, all filled row 3 to the last row:
   - **Z** ("L2 Due Date"): `=IF(X{row}="","",WORKDAY.INTL(X{row},1,11,a))`,
     with `number_format = "dd-mmm-yy"` (same date format as the N/"Due Date"
     column). **Depends on X (`Form submisison - L2`), not Y** — a deliberate
     change from an earlier version that used Y (`L2 check addition
     (Oldest)`); don't "fix" this back to Y. `a` is a genuine Excel
     workbook-scoped named range (`wb.defined_names["a"]`, not a Python
     variable) pointing at `'Holiday List'!$A$2:$A${last_row}` — see the
     Holiday List gotcha below. Verified end-to-end against the real
     `Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx`: sheet created, named range resolves, formula
     text matches the given spec exactly, survives a save/reload round-trip
     and the styling pass.
   - **AA** ("L2 TAT"):
     `=IFERROR(IF(Z{row}="","",IF(INT(U{row})<=INT(Z{row}),"IT","OT")),"")` (U
     = `L2 Report sent date`). Wrapped in `IFERROR` and compares `INT(...)`
     of both sides (strips any time-of-day component before comparing pure
     dates) — both deliberate per the given spec, don't simplify back to the
     bare `IF(U{row}<=Z{row},...)` form.
   - **AB** ("L2 check status"):
     `=IF(AND(X{row}<>"",Y{row}<>"",U{row}<>""),K{row},IF(AND(X{row}="",Y{row}<>""),"Pending at candidate",IF(X{row}<>"","Work In Progress","")))`
     (K = `Case Status`, used as the fallback status value). **Checks Y ("L2
     check addition (Oldest)"), not Z** — changed at explicit user request
     (previously checked Z, "L2 Due Date"). This isn't a cosmetic swap: Z is
     *derived from* X (`Z = IF(X="","",WORKDAY.INTL(X,1,11,a))`, i.e. Z is
     blank exactly when X is blank), so the old `AND(X="",Z<>"")` → `"Pending
     at candidate"` branch could mathematically never fire — it was dead
     code. Y comes from an independent VLOOKUP against `'L2 check addition
     (Oldest)'!B:T` (added by `add_l2_check_addition_oldest_sheet`),
     unrelated to X, so a case can genuinely have Y populated (reopened) with
     X still blank (no form resubmission yet) — the "Pending at candidate"
     branch is real and reachable now. Don't "fix" this back to Z thinking
     it's restoring consistency with AA (AA legitimately still uses Z — see
     above — only AB changed).

## `add_status_view_sheet()` — the "Status view" report

Builds a `Status view` sheet (moved to tab index 0) holding **genuine, live
Excel PivotTables** — not a static openpyxl-computed imitation (an earlier
version of this function built one; it was replaced after the user
explicitly asked for real pivot tables). Seven of them: **Received cases**,
**Case status**, **L1 reports sent severity**, **L1 TAT**, **L2 reports TAT**
(by Case status), **L2 check status**, and **L2 reports sent severity**.
Modeled on the "Status view" tab of a separate reference workbook,
`Mphasis_Limited_Pre_Offer_checkwise Dashboard V1.1.xlsx` — that file is
**only a layout reference**, never opened/read/linked by this app. Every
pivot here is built fresh off `Mphasis_Limited_Pre_Offer_checkwise
Dashboard.xlsx` each run, and all 7 (including L2 reports sent severity) are
built off the single `Pre_Offer_checkwise`-sourced `cache` — no pivot reads a
separate audit sheet as its own `PivotCache` source. **L2 reports sent
severity** was changed at explicit user request to read `Pre_Offer_checkwise`
directly — `row_field="L2 Report sent date"`, `col_field="L2 Report sent
severity"`, `data_field="ARS Number"` (the same `dd-mmm-yy` row field +
severity column field + ARS Number count pattern as **L1 reports sent
severity**) — instead of the earlier design that built a separate
`PivotCache` off the `L2 Report Sent` audit sheet (`report_sent_on` /
`report_severity` / `case_ars_no`) with a Year>Month>Day row hierarchy. Since
`L2 Report sent date`/`L2 Report sent severity` are original tracker columns
present in `Pre_Offer_checkwise` from the start (see
`update_case_status_from_l2_report_sent` / `add_l2_report_sent_sheet`), this
pivot is no longer gated on the `L2 Report Sent` sheet's existence the way
**L2 reports TAT**/**L2 check status** still are.

- **Why this drives Excel via COM instead of openpyxl**: openpyxl cannot
  create real `PivotTable`/`PivotCache` objects at all — there's no API for
  it. This function opens a **hidden Excel instance via `win32com`** (the
  same automation technique already used for Outlook) and builds the pivots
  through the real Excel object model. This incidentally also solves a
  second problem for free: Pre_Offer_checkwise's N/O/U/V/W/X/Y/Z/AA/AB
  columns are all formulas that openpyxl never evaluates (they'd read back
  as unevaluated/`None`), but a real Excel session opening the file and
  running `Application.CalculateFullRebuild()` resolves all of them for
  real — confirmed live (`O3` → `"IT"`, `Z3` → a real date, `AA3` → `"OT"`,
  `AB3` → `"Completed"`) — so the pivots can aggregate directly on
  Pre_Offer_checkwise's own columns with zero Python reimplementation of
  that formula logic.
- **`_XL_*` constants are hardcoded integers, not named enums.** Late-bound
  `win32com.client.Dispatch("Excel.Application")` (used here, same as the
  Outlook integration) has no access to Excel's named VBA constants
  (`xlRowField`, `xlCount`, etc.) — these are the stable, documented enum
  values (`xlRowField=1`, `xlColumnField=2`, `xlCount=-4112`, `xlDatabase=1`,
  `xlUp=-4162`, `xlToLeft=-4159`).
- **Gotcha: Excel's native date-grouping (`Range.Group`) reliably FAILS via
  this COM automation path** — reproduced directly, repeatedly, even on a
  trivial freshly-hand-built numeric field with zero connection to this
  app's data (a plain `Workbooks.Add()` with 60 rows, grouping by 10), with
  every combination of argument style tried (positional/named/tuple/explicit
  Start-End dates) failing identically with the same generic "Group method of
  Range class failed" COM error. This isn't a data-quality issue — it's
  something about `Range.Group` via this specific binding/environment. This
  is why every date row field in `Status view` (`Case Received Date`, `L1
  Report Sent Date`, `L2 Report sent date`) is a plain row field with a
  `row_number_format` rather than a native Year>Month>Day grouping.
  `_add_pivot_table`'s `extra_row_fields` param still exists as the
  workaround mechanism for a caller that needs an explicit multi-level row
  hierarchy without `Range.Group` — it was originally built for exactly this
  on the L2 reports sent severity pivot (two helper columns, `Year`/`Month`,
  appended past the end of the source sheet via `=YEAR(RC{col})` /
  `=DATE(YEAR(RC{col}),MONTH(RC{col}),1)` formulas, added as extra outer row
  fields) — but no pivot currently uses it, since that pivot was changed to
  read `Pre_Offer_checkwise` directly (see above) without a nested hierarchy.
  If `Range.Group` ever needs revisiting (e.g. after an Excel/Office update),
  re-verify with a from-scratch repro before assuming it's fixed.
- **Gotcha (historical — no pivot currently exercises this): row-field
  nesting order is controlled by `PivotField.Position`, not by the order
  `.Orientation` is assigned.** Confirmed live, back when the L2 reports sent
  severity pivot used the `Year`/`Month` helper-column workaround above:
  setting `Orientation` on `Year`, then `Month`, then the date field, in that
  order, still nested them date-field-first (Excel orders by the field's
  underlying column position in the source range, not assignment/call order)
  until each field's `.Position` was set explicitly. Relevant again if
  `extra_row_fields` is ever used by a future pivot.
- **Gotcha: `PivotField.NumberFormat` and blank-item hiding must happen
  AFTER `AddDataField()` + `RefreshTable()`, not right after `Orientation`.**
  Setting `NumberFormat` immediately after `Orientation` intermittently threw
  "Unable to set the NumberFormat property of the PivotField class" — moving
  it to a second pass after the table has a data field and has been
  refreshed at least once fixed it for most fields. `L1 Report Sent Date`'s
  `NumberFormat` (shared between the L1 severity and L1 TAT pivots, both
  built off the same `PivotCache`) still intermittently fails even post-
  refresh — root cause not fully pinned down (possibly a shared-cache-field
  formatting conflict between the two pivot tables using it), so `NumberFormat`
  assignment is wrapped in a bare `try/except: pass` and treated as
  best-effort cosmetic polish, never fatal to the report. When it fails, the
  affected date column just falls back to whatever format the source cells
  already carried (still a valid, readable date — just not `dd-mmm-yy`).
- **Gotcha: a `PivotItem`'s blank/no-value category is sometimes named the
  literal empty string `""` and sometimes the literal text `"(blank)"`,
  inconsistently across fields** — confirmed both forms live on the same
  workbook (`L2 TAT`/`L2 check status` used `""`, nothing here has shown a
  real `"(blank)"` case yet, but both are checked in
  `_hide_blank_pivot_item`). Don't assume one name and drop the other check.
- **Gotcha: `L2 reports TAT` and `L2 check status` are skipped entirely —
  not built empty — when `Form submisison - L2` / `L2 Report Sent` don't
  exist yet** (i.e. right after "Clean Latest File" alone, before the MIS
  query chain has run). Building them anyway throws "`PivotFields` method of
  `PivotTable` class failed" — Pre_Offer_checkwise's `L2 TAT`/`L2 check
  status` formulas reference those sheets by name directly, and Excel
  resolves the reference to a real error value when the sheet doesn't exist,
  which breaks `PivotFields()` on that column outright (not just
  IFERROR-catchable at the formula level, at the pivot-field level). The MIS
  button handler calls `add_status_view_sheet()` again at the end of its
  chain, once both sheets exist, which rebuilds the sheet with all 7. **`L2
  reports sent severity` is not gated this way** — since it now reads
  `Pre_Offer_checkwise`'s own `L2 Report sent date`/`L2 Report sent severity`
  columns (present in the tracker from the start, not sheet-name formula
  lookups), it builds every run, same as the L1 pivots — just with
  whatever's currently in those two columns (stale/empty until the MIS Sent
  Cases query populates them via `add_l2_report_sent_sheet`).
- **Gotcha: `Application.Quit()` unreliably leaves `EXCEL.EXE` resident when
  `Visible=False`.** Reproduced repeatedly, including with `Workbooks.Count
  == 0` at the moment `Quit()` was called (so it isn't a "forgot to close a
  workbook" issue) — this is a known limitation of hidden Excel COM
  automation, not a bug in the cleanup code. Worked around by capturing the
  process ID via `win32process.GetWindowThreadProcessId(excel.Hwnd)` right
  after `Dispatch()`, then force-`taskkill`ing it in the `finally` block if
  it's still alive ~2 seconds after `Quit()`. Don't remove this thinking
  `Quit()` alone is sufficient — repeated runs will accumulate zombie
  `EXCEL.EXE` processes without it.
- **Verified safe: the PivotTables survive `apply_professional_styling()`'s
  later openpyxl `load_workbook()` + `wb.save()`.** openpyxl doesn't natively
  model PivotTables, so there was real risk that round-tripping the file
  through it would corrupt or silently drop them — confirmed directly
  instead of assumed: built all 7 pivots via COM, ran the exact
  `load_workbook()`/`.save()` sequence `apply_professional_styling()` uses,
  reopened via COM, and confirmed all 7 were still present with correct
  dimensions AND still successfully `RefreshTable()`-able. If openpyxl's
  pivot-handling ever changes (version bump), re-verify this before trusting
  it still holds.
- **Column layout is computed dynamically, not hardcoded to the reference
  file's column letters.** Each pivot table gets its own column band, placed
  2 columns to the right of the previous pivot's actual rightmost occupied
  column (via `pt.TableRange1.Column + pt.TableRange1.Columns.Count`) — never
  a fixed offset. This is what "realign to get rid of overlapping" turned
  into: real PivotTables refuse to overlap on refresh, and the number of
  distinct categories (hence table width) varies run to run with the data,
  so a fixed layout copied from the reference file's column letters would
  eventually collide.
- **Rebuilt from scratch on every run** (existing `Status view` sheet is
  deleted first), same idempotent pattern as the other `add_*_sheet`
  functions.
- **Cosmetic polish, added at explicit user request** ("beautify it and make
  professional") — all done via COM in `_add_pivot_table`/`add_status_view_sheet`
  since openpyxl can't touch PivotTable formatting at all:
  - Each pivot gets a **merged, navy-fill/white-bold title banner** spanning
    its actual measured width (`_STATUS_VIEW_TITLE_FILL`/`_STATUS_VIEW_TITLE_FONT_COLOR`,
    matching the same `1F4E78`/`FFFFFF` theme `HEADER_FILL`/`HEADER_FONT` use
    everywhere else in this workbook) — computed via `_rgb_hex_to_ole_bgr()`,
    since Excel COM color properties (`Interior.Color`, `Font.Color`,
    `Tab.Color`) take an OLE COLOR (`0x00BBGGRR`), not RGB hex; don't assign
    an RGB hex int directly to one of these, it'll render the wrong color.
  - `pt.TableStyle2 = _PIVOT_TABLE_STYLE` (`"PivotStyleMedium9"`) applied to
    every pivot for consistent banded-row styling — wrapped in
    `try/except: pass` (cosmetic, never fatal).
  - `status_ws.UsedRange.Columns.AutoFit()` once at the end, after all 7
    pivots exist, so labels like `"Major Discrepancy"` or `"Pending at
    candidate"` aren't clipped.
  - The sheet tab color (`_STATUS_VIEW_TAB_COLOR`, green `2E7D32`) now goes
    through the same `_rgb_hex_to_ole_bgr()` helper instead of a hand-rolled
    bit-shuffle — same result, just clearer to read/maintain.

## Key gotchas

- **Multiple Outlook accounts: `GetDefaultFolder(6)` only searches the DEFAULT
  account's Inbox.** On the actual deployment machine, Outlook has two
  top-level accounts (`client_servicing@authbridge.com` and `Google Workspace -
  jay.chaudhary@authbridge.com`), and the tracker email lands in the latter —
  which may not be Outlook's "default" one. `get_target_inbox(namespace,
  account_name)` looks up the matching top-level folder by substring on
  `folder.Name` (e.g. `"jay.chaudhary@authbridge.com"` matches the `"Google
  Workspace - ..."` folder) and uses *its* Inbox instead. The "Outlook account"
  UI field feeds this — leaving it blank falls back to `GetDefaultFolder(6)`.
  If someone reports "times out with 'No email found'" despite the email
  visibly sitting in their Inbox, check this first, not the subject-matching
  logic.
- **Email-scan depth**: `download_latest_from_email()` only checks the N most
  recent items per poll (`max_check`, currently 300, was 51). A busy/shared
  inbox can push the target email past a too-small N well before the search
  timeout — this was the second half of the "times out despite email existing"
  bug above (email was at position 81 in a 998-item inbox getting ~13
  emails/hour). If this recurs on an even busier mailbox, raise `max_check`
  further rather than assuming it's a subject-matching issue.
- **`DOWNLOAD_DIR`** is hardcoded to `C:\Gen AI\Mphasis\Downloads`. The final
  output file name (`Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx`,
  renamed from `cleaned_tracker.xlsx`) lives in the `CLEANED_FILE_NAME`
  constant — every function references it through that constant rather than
  a hardcoded literal, so renaming it again only means changing one line.
  `get_latest_excel_file()` explicitly excludes `CLEANED_FILE_NAME` (matched
  by stem, and any `~$`-prefixed Excel lock file) from consideration —
  without this exclusion it would treat its own output as the "latest"
  source on a second run.
- **`Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx` open in Excel blocks saving.** Any clean/export run
  will fail with `PermissionError: [Errno 13]` if the file is open elsewhere —
  look for a `~$Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx` lock file in `DOWNLOAD_DIR` to confirm.
- **Target sheet**: the source workbook has multiple sheets (`Pre_Offer_checkwise`,
  `Summary`, ...). Cleaning only ever operates on the sheet matched by
  `"pre_offer_checkwise" in name.lower()`.
- **Header row is row 2, not row 1.** Row 1 is a title line (e.g.
  `"Mphasis_Limited as on <date>"`). All header lookups in
  `clean_latest_excel_file()` must scan `ws[2]`. This was the root cause of a bug
  where column-deletion silently did nothing — get this wrong again and the same
  symptom will reappear.
- **Column deletion is header-name based, not position based.** Column letters
  shift between tracker exports, so `clean_latest_excel_file()` matches on exact
  (case-insensitive, stripped) header text in a set (`headers_to_delete`) and
  loops, re-scanning after each delete, until no more matches are found — this
  also handles duplicate headers appearing more than once in the row (e.g.
  "National Identity Check (2)" has historically appeared at more than one
  column position in a given export).
- Current `headers_to_delete` set: `National Identity Check (2)`,
  `Criminal Records Verification (2)` through `(9)`.
- **`HOLIDAY_DATES_RAW` (the static holiday calendar behind the `a` named
  range) has a block of 10 entries with no year** (`01-Jan` through
  `25-Dec`, sitting between `25-Dec-18` and `21-Oct-19` in the given list —
  every other entry has an explicit year, so this was a source formatting
  glitch, not intentional). `_parse_holiday_date()` infers these as **2019**
  based on chronological position (`_HOLIDAY_YEAR_LESS_INFERRED_YEAR = 2019`)
  — flagged to the user at the time rather than silently assumed, but never
  independently re-confirmed against an actual holiday calendar. All 216
  entries parse successfully (verified — 0 unparsed). If this list is ever
  regenerated from a fresh source, re-check that block specifically rather
  than trusting the inference held over.
- **DB credentials live in `.env`** (`DB_HOST`, `DB_PORT`, `DB_NAME`, `DB_USER`,
  `DB_PASSWORD`), never hardcoded in `mphasis.py` — loaded via `python-dotenv`'s
  `load_dotenv()` at import time, read through `os.getenv(...)`. `.env` is
  gitignored; don't commit it, and don't move these back to hardcoded
  constants. `DB_PORT` defaults to `3306` if unset. This replaced an earlier,
  explicitly-approved-at-the-time choice to hardcode the (now-defunct) MIS
  website's login credentials directly in `mphasis.py` — that tradeoff no
  longer applies now that there's a real DB connection with its own
  credentials, which went straight into `.env` from the start.
- **`check_id1` is swapped for a parameterized IN-list via plain
  `str.replace()`, never f-string/`%`-interpolation of the ARS numbers
  themselves.** `_execute_ars_query()` does
  `sql_template.replace("check_id1", ",".join(["%s"] * len(ars_numbers)))`
  then `cursor.execute(query, ars_numbers)` — the placeholder *text* is
  substituted, but the actual ARS number *values* always travel as bound
  parameters through `mysql-connector-python`, not string-formatted into the
  query. Don't "simplify" this into building the IN-list by joining quoted
  ARS numbers into the SQL text directly — that would reintroduce SQL
  injection risk for no benefit.
- **DB result rows need type handling that CSV-parsed rows never did** — see
  "Two type-handling differences from the old CSV-parsing path" earlier in
  this doc for the full detail (`_none_first_sort_key()` for sorting columns
  that may be a real `datetime` or `None`; `_excel_safe()` to convert
  `decimal.Decimal` to `float` before writing into openpyxl). Both were
  discovered and fixed during this rewrite, not carried over from the CSV
  era — the CSV path never had this problem because every CSV cell was
  already a plain string.
- **Historical: Selenium → Playwright → direct DB query.** MIS data
  retrieval went through three implementations: Selenium/ChromeDriver (slow,
  fragile Export-click hangs), then Playwright (fixed the hangs via
  `page.expect_download()`, but still needed `channel="chrome"` to dodge a
  network block on the bundled Chromium binary, an asyncio event-loop-policy
  reset to coexist with Streamlit/Tornado, and a workaround for the MIS
  site's Query dropdown silently resetting after a Data-Time-Slab
  reselection), and finally this direct DB connection, which removes the
  browser — and every one of those browser-specific workarounds — entirely.
  None of that Selenium/Playwright code exists in `mphasis.py` anymore; if
  DB access ever becomes unavailable and browser automation needs reviving,
  treat this as a fresh implementation rather than assuming the old
  Playwright code (no longer present) can just be restored.
- **`Report Status` in the Sent Cases query result is lowercase (`"sent"`),
  not `"Sent"`** — it comes straight from the query's own `WHEN 5 THEN
  'sent'` CASE branch (see `SENT_CASES_QUERY`), so this isn't a data quirk to
  work around so much as a fact about the query itself. `add_l2_report_sent_sheet()`
  compares `.strip().lower()` against `REPORT_SENT_STATUS.lower()` /
  `REPORT_SENT_TYPE.lower()` for exactly this reason — don't revert to exact
  `==` matching on these two columns. Confirmed at production scale in this
  doc's "MIS data retrieval" section above (3,551 of 6,964 Sent Cases rows
  matched `sent`+`Additional`).

## Debugging header-matching issues

If a "clean" column doesn't disappear, don't guess — inspect the actual
downloaded file directly:

```python
from openpyxl import load_workbook
wb = load_workbook("Downloads/<file>.xlsx", data_only=True)
ws = wb["Pre_Offer_checkwise"]
for cell in ws[2]:
    if cell.value:
        print(cell.column, repr(cell.value))
```

This is how the row-1-vs-row-2 header bug was found — the hardcoded assumption
didn't match reality until the real file was opened and inspected. The same
"open the real file, don't guess" approach was used to discover the MIS site's
actual dropdown flow (Host → Database → Data Time Slab → Query) and the exact
field/button XPaths before writing any browser-automation code against them.

## Historical: testing the Selenium/Playwright MIS automation (superseded)

Before MIS data retrieval moved to a direct DB connection (see "MIS data
retrieval — direct DB query" above), this app drove the MIS export-query
website via browser automation and was tested against both the live site and
a local mock server replicating its login form, async dropdown rebuilds, and
slow `/export` endpoint. None of that browser-automation code, nor the
XPath/download-navigation quirks it worked around, exists in `mphasis.py`
anymore — testing this app now means testing DB queries and Excel output, not
browser mechanics. Kept as a one-line pointer only in case browser automation
of the MIS website is ever revived from scratch in the future.
