# Mphasis Tracker

Single-file Streamlit app (`mphasis.py`) that downloads the latest Mphasis ARS
tracker Excel file from Outlook, cleans and annotates it, then drives the
AuthBridge MIS export-query tool to pull a supplementary "Major Discrepancy"
dataset and merge it back in.

## Run

```
streamlit run mphasis.py
```

Requires Outlook to be running on the same Windows machine (uses `win32com.client`
COM automation, not the Graph API). Also requires the `playwright` package and a
real installed Chrome browser for the MIS automation step — `run_ars_query()`
launches via `channel="chrome"` rather than Playwright's bundled Chromium (see
gotcha below for why); `playwright install chromium` alone is not sufficient.
Also requires a real installed copy of Excel on the same machine —
`add_status_view_sheet()` drives it via COM automation (a hidden instance) to
build genuine, live PivotTables, since openpyxl cannot create real
PivotTable objects at all.

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
4. **Run ARS Query in MIS** — one button handler chains: `run_ars_query` (login +
   Advance Tracker) → `export_and_download` → `add_red_remarks_sheet` →
   `select_query` + `fill_ars_number_field` (Case History, same session) →
   `export_and_download` again → `add_form_submission_l2_sheet` →
   `add_l2_check_addition_oldest_sheet` → `select_query` + `fill_ars_number_field`
   (Sent Cases, same session) → `export_and_download` again →
   `add_l2_report_sent_sheet` → `update_case_status_from_l2_report_sent` →
   `add_l2_summary_columns` → `add_status_view_sheet` → `apply_professional_styling`.
   See below. (`clean_latest_excel_file`, step 3, also calls `_write_holiday_list_sheet()`
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

## Step 4 in detail: MIS export-query automation

As of this rewrite, this step is driven by **Playwright**, not Selenium (see
the "Selenium → Playwright" gotcha below for why, and what changed).

`run_ars_query(data_time_slab_keyword, query_name)` drives Playwright/Chromium to:
1. Log into `https://mis.authbridge.com/export_query/login.php` with
   `MIS_USERNAME` / `MIS_PASSWORD` (hardcoded constants — see gotcha below).
2. Select Host → `MIS_HOST` ("Bridge Live"), Database → `MIS_DATABASE`
   ("Bridge Live"). Both are currently hardcoded; only the query type/name vary
   per call.
3. Call `select_query(page, data_time_slab_keyword, query_name)`, which
   selects **Data Time Slab** to the option matching `data_time_slab_keyword`
   as a case-insensitive substring (e.g. `"Case Query"` matches the actual
   option text `"case query - Bridge"`) and **Query** to `query_name` exactly
   (e.g. `"Advance Tracker"`, `"Case History"`). This reveals an `ARS No*`
   field — `input[name="check_id1"]`, i.e.
   `//*[@id='date']/tbody/tr[2]/td[2]/input`.

`select_query()` and `fill_ars_number_field()` are standalone (not just internal
to `run_ars_query`) specifically so a **second query can reuse the same logged-in
browser session** instead of opening another Chrome window and logging in again
— see the button handler flow below.

Every step below that reads an MIS export (`add_red_remarks_sheet`,
`add_form_submission_l2_sheet`, `add_l2_check_addition_oldest_sheet`,
`add_l2_report_sent_sheet`) goes through the shared `_read_export_csv_rows()`
helper first, which transparently handles both a plain `.csv` and a
`.csv.zip` download (Chrome/the site can hand back either) and returns the
parsed rows including the header — none of those functions parse the file
format themselves.

The button handler:
1. Builds the ARS number list **directly from `Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx` column C**
   via `get_ars_number_list()` — not by reading the AC column's formula (openpyxl
   never evaluates formulas, so AC would just return unevaluated formula text).
   The generated string is equivalent to what AC would show once opened in Excel.
2. `run_ars_query("Case Query", "Advance Tracker")` → `fill_ars_number_field()`
   pastes it into the ARS No field.
3. `export_and_download()` clicks the Export button
   (`/html/body/div/div/div[2]/form/input`) inside a `page.expect_download()`
   block and saves the resulting `Download` object into `MIS_EXPORT_DIR`
   (`Downloads/MIS_Exports/`, kept separate from tracker files so it never gets
   picked up as the "latest" source file). This listens for the browser's
   actual download event directly, rather than polling the filesystem for a
   new file — see the gotcha below for why that distinction mattered.
4. `add_red_remarks_sheet(downloaded_file)`:
   - reads the exported file (plain `.csv` or `.csv.zip`, both handled)
   - filters rows to `check_severity` in `RED_REMARKS_SEVERITIES` (`Major
     Discrepancy`, `Amber`, `Minor Discrepancy`) **AND** `Check_unique_name`
     in `RED_REMARKS_CHECK_NAMES` (`UAN Check for Undisclosed Employment`,
     `Dual Employment Verification via Form 26AS`, `Criminal Records
     Verification`, `National Identity Check`, `India Court Record Check
     through Law Firm`) — both exact match, both required
   - groups the filtered rows by `case_ars_no` and joins each ARS's
     `closure_comments` with `", "` into **one row per unique ARS** (an ARS
     can have multiple matching checks, e.g. both a Dual Employment and a UAN
     check) — this is a real behavior change from the original "keep every
     matching row as-is" version; verified to reproduce the exact expected
     combined-comment text against a real export for 16 sample ARS numbers
     (e.g. ARS `3055-016865` → `"Possible dual employment is found....,
     Verification could not be possible as the EPFO portal is not
     working...."`, joined in row-encounter order, not sorted)
   - writes `case_ars_no` → combined `closure_comments` into a new `red
     remarks` sheet in `Mphasis_Limited_Pre_Offer_checkwise Dashboard.xlsx` (column B = `case_ars_no`,
     column C = combined `closure_comments`; replacing the sheet if it
     already exists) — **not** the full ~75-column raw rows the original
     version kept for audit purposes; this sheet is now a 3-column (A blank,
     B, C) aggregation, not a raw data dump
   - fills `Pre_Offer_checkwise!W3:W{last row}` — the tracker's existing
     `Red Remarks` header column — with
     `=IFERROR(VLOOKUP(C{row},'red remarks'!B:AB,2,0),"")`, which looks up the
     ARS number against `red remarks` column B (`case_ars_no`) and returns
     column C (the combined `closure_comments`, the 2nd column of the B:AB
     range — the `AB` upper bound is arbitrary headroom, nothing needs to
     actually exist past column C for this to work)
5. On the **same** `page`, `select_query(page, "Case Query", "Case
   History")` (option `//*[@id='...']/select/option[15]` in the Query dropdown,
   verified live) + `fill_ars_number_field()` re-pastes the same ARS number list,
   then `export_and_download()` downloads `Case History.csv.zip` (much larger
   than the Advance Tracker export, ~2.3MB vs ~160KB, since it's full case
   history detail rather than one row per check).
6. `add_form_submission_l2_sheet(case_history_file)`:
   - filters the Case History rows to `ACTION_COMMENTS == FORM_SUBMISSION_L2_COMMENT`
     (`"Check moved to WIP. Antecedents populated from iBridge candidate
     submission."`), sorts by `ACTION_TAKEN_ON` oldest-first (plain string sort —
     safe since the format is zero-padded `YYYY-MM-DD HH:MM:SS`)
   - writes the result into a new `Form submisison - L2` sheet (that exact
     spelling/typo, kept verbatim per explicit instruction — don't "fix" it)
   - fills `Pre_Offer_checkwise!X2` with that same sheet name as the header, and
     `X3:X{last row}` with
     `=IFERROR(VLOOKUP(C{row},'Form submisison - L2'!B:T,12,0),"")` (column B =
     `case_ars_no`, the 12th column of B:T = `ACTION_TAKEN_ON`)
7. `add_l2_check_addition_oldest_sheet(case_history_file)`:
   - filters the same Case History rows to `ACTION_TAKEN == CASE_REOPENED_ACTION`
     (`"New Status - Case Reopened"`), same oldest-first sort by `ACTION_TAKEN_ON`
   - writes the result into a new `L2 check addition (Oldest)` sheet
   - fills `Pre_Offer_checkwise!Y2` with that same sheet name as the header, and
     `Y3:Y{last row}` with
     `=IFERROR(VLOOKUP(C{row},'L2 check addition (Oldest)'!B:T,12,0),"")` (same
     B:T/column-12 alignment as the `Form submisison - L2` lookup — column B =
     `case_ars_no`, the 12th column of B:T = `ACTION_TAKEN_ON`)
8. On the **same** `page`, `select_query(page, "Case Query", "Sent
   Cases")` + `fill_ars_number_field()` re-pastes the same ARS number list,
   then `export_and_download()` downloads the Sent Cases export.
9. `add_l2_report_sent_sheet(sent_cases_file)`:
   - filters rows to `Report Status == "Sent"` AND `Report Type == "Additional"`
     (both exact match), sorts by `report_sent_on` oldest-first
   - writes the full matching rows into a new `L2 Report Sent` sheet
     (**column order matches the raw Sent Cases CSV as-is** — `case_ars_no`
     lands at column **F**, index 5, not B; confirmed live — anything reading
     this sheet must look up columns by header name, never assume B), then
     reparses just that sheet's `report_sent_on` column from the exported
     datetime string into a real `date` value (needed so the destination
     `dd-mmm-yy` number format below actually applies — a plain text string
     ignores number formats)
   - **clears** `Pre_Offer_checkwise`'s existing `L2 Report sent date` / `L2
     Report sent severity` column values first, then refills them via
     `INDEX/MATCH` (not `VLOOKUP`) against `L2 Report Sent`, matching on
     `case_ars_no` vs. `Pre_Offer_checkwise!C{row}`. INDEX/MATCH was chosen
     over VLOOKUP here specifically because `case_ars_no` is not known to sit
     to the left of `report_sent_on`/`report_severity` in that sheet (a hard
     requirement for VLOOKUP, but not for INDEX/MATCH) — confirmed true:
     `case_ars_no` (F) actually sits *right* of neither in this export's
     column order, so this wasn't just defensive. Both destination columns
     are located by **header name** via `_find_header_col_letter()` (scanning
     row 2) rather than hardcoded letters, for the same reason — see the
     debugging section below before assuming this is wired to the wrong
     columns.
10. `update_case_status_from_l2_report_sent()`:
    - for `Pre_Offer_checkwise` rows whose `Case Status` (K) is currently
      `Insufficient` or `On Hold` (both confirmed exact-match strings —
      the only other real values seen are `Completed` and `Work In
      Progress`), overwrites K to `Completed` if that row's ARS number
      appears in the `L2 Report Sent` sheet's `case_ars_no` column (looked
      up by header name — see step 9's column-order gotcha above; a hardcoded
      `B` here was a real bug caught before it shipped), otherwise leaves K
      untouched
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
11. `add_l2_summary_columns()` — runs last, since it depends on X (step 6),
    Y (step 7), U/"L2 Report sent date" (step 9), and the `a` named range
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
     (Oldest)'!B:T` (added by `add_l2_check_addition_oldest_sheet`, step 7),
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
- **MIS credentials are hardcoded in plain text** (`MIS_USERNAME`,
  `MIS_PASSWORD` near the top of `mphasis.py`) — an explicit user choice over
  env vars/UI input. Don't commit this file anywhere shared without stripping
  them, and don't "fix" this to env vars without asking first.
- **`run_ars_query()` launches `channel="chrome"` (the real installed Chrome),
  not Playwright's bundled Chromium — this is load-bearing, not cosmetic.**
  On the deployment network, the bundled Chromium got `net::ERR_CONNECTION_CLOSED`
  hitting `mis.authbridge.com` every time, while `curl` and a real Chrome
  instance both reached it fine — some corporate proxy/security layer
  evidently allows the recognized Chrome executable through while blocking
  the unrecognized bundled one. Confirmed directly: reproduced the failure
  with the bundled binary, then confirmed `channel="chrome"` fixes it,
  end-to-end, against the live site. If MIS navigation starts failing with
  `ERR_CONNECTION_CLOSED` (or similar) again, don't assume the site is down —
  check whether this got reverted first.
- **`run_ars_query()` resets the asyncio event loop policy to
  `WindowsProactorEventLoopPolicy` before starting Playwright — this is also
  load-bearing.** Streamlit is built on Tornado, which forces the
  process-wide asyncio policy to `WindowsSelectorEventLoopPolicy` on Windows;
  `SelectorEventLoop` can't spawn subprocesses on Windows, and Playwright's
  sync API needs to spawn its driver subprocess on startup, so calling
  `sync_playwright().start()` from inside a running Streamlit app raised
  `NotImplementedError` deep inside `asyncio`'s subprocess machinery — every
  single time, with no exception message text (`str(e)` was empty; only
  `type(e).__name__` showed anything, which is why `except Exception as e:`
  handlers in this app show `str(e) or type(e).__name__`, not just
  `str(e)`). This is exactly why the *identical* code always worked when run
  as a standalone script (default Windows policy is already Proactor there)
  but always failed through the actual Streamlit UI — a hard-to-track-down
  discrepancy that cost significant back-and-forth before being root-caused.
  Confirmed by reproducing the exact traceback outside Streamlit just by
  forcing the Selector policy manually, then confirming the reset fixes it.
  Don't remove this thinking it's unnecessary defensive code.
- **Selenium → Playwright rewrite.** The MIS automation originally used
  Selenium/ChromeDriver, and was rewritten to use Playwright after repeated,
  hard-to-diagnose hangs on the Export click (a raw `HTTPConnectionPool(...):
  Read timed out (read timeout=120)` from Selenium's own HTTP client — no
  indication it happened at Export, no catchable exception). Root cause:
  Selenium's `.click()` on the Export button (a `<form>` submit) blocks until
  the browser reports the resulting navigation complete; a slow server-side
  report generation could hang that single `.click()` call for minutes,
  *before* `export_and_download()`'s own polling loop even started counting
  (it only began after `.click()` returned). Capping Selenium's page-load
  timeout and catching the resulting `TimeoutException` worked around one
  instance of this, but the underlying design — click, then hope, then poll
  the filesystem — was fragile by construction. Playwright's
  `page.expect_download()` sidesteps the whole problem: it listens for the
  browser's actual download event directly, decoupled from whatever the
  page's navigation/load state is doing, so a slow server response no longer
  risks hanging the click itself. If MIS automation issues come back, verify
  first whether they're actually Selenium-shaped (they shouldn't be anymore —
  there's no more ChromeDriver, no more ports, no more raw HTTP timeouts) before
  assuming this class of bug has resurfaced.
- **No more staleness handling needed for dropdowns/fields.** Selenium's
  `StaleElementReferenceException` dance (`_select_dropdown_option()`
  re-locating on every retry, `fill_ars_number_field()` retrying the whole
  locate→clear→send_keys sequence) doesn't apply to Playwright: locators
  re-resolve the DOM by selector on every action rather than holding a live
  element handle, and `page.fill()` / `page.select_option()` auto-wait for the
  target to become actionable. `_select_dropdown_option()` still polls in a
  loop, but only because the Data Time Slab / Query rebuild is *asynchronous*
  (the matching `<option>` genuinely doesn't exist for ~400ms after the
  triggering `onchange`) — not to work around staleness. Don't reintroduce a
  staleness-retry pattern here; it was solving a Selenium-specific problem.
- **`table#date` (the ARS No field's container) rebuilds when the *Query*
  dropdown changes, not when Data Time Slab changes** — confirmed via a local
  mock-server test built specifically to validate the Playwright rewrite (see
  below). Don't assume Data Time Slab changes are what invalidate the ARS
  field.
- **Re-selecting Data Time Slab (`access_time`) — even to its already-selected
  value — silently resets the Query dropdown (`csv_query`) back to blank
  shortly afterward, without rebuilding its option list.** This broke
  switching queries on a reused session (e.g. Advance Tracker → Case
  History): `select_query()` re-selects `access_time` first (needed for the
  *first* query of a session, harmless-looking to repeat for later ones), then
  immediately selects `csv_query` — and that immediate selection was getting
  silently wiped a moment later. `select_option()` not raising an exception
  does **not** mean the selection survived. Confirmed live by direct
  reproduction: selecting `csv_query` right after re-selecting `access_time`
  reverted to blank within ~300ms and stayed blank; doing the identical
  `select_option()` call after a several-second pause stuck reliably.
  Symptom in the UI was maximally confusing — no exception, no error on the
  *first* query of a session (Advance Tracker), and the *second* query's
  Export click silently triggered the site's own `alert('Please select
  Query.')` (auto-dismissed by Playwright by default, so nothing surfaced it)
  before `expect_download()` timed out with the generic "Export clicked but
  no download completed" message — indistinguishable from a genuinely slow
  server response unless you go looking at what's actually selected.
  Fixed in `_select_dropdown_option()`: after `select_option()`, it now
  re-reads the select's actually-selected option text and retries the whole
  selection if it doesn't match what was just set, rather than trusting
  `select_option()` not raising as proof of success. This applies to *every*
  dropdown selection now (hostname/database/access_time/csv_query), not just
  csv_query — don't special-case it back down to just one dropdown; the
  general "verify, don't just trust" pattern is the actual fix, and the exact
  reset trigger/timing on the site's side is still unconfirmed beyond this
  one reproduction.
- **Playwright browsers are never closed either** (`.close()` / `.stop()` are
  never called, mirroring the old "never call `driver.quit()`" intentional
  behavior — the browser stays open per an explicit prior user request, so it
  can be reviewed / reused for later queries on the same session).
  `run_ars_query()` calls `sync_playwright().start()` directly rather than
  using the `with sync_playwright() as p:` context-manager form specifically
  *because* exiting that `with` block would tear down the driver connection
  and close the browser — don't refactor this into a `with` block, it would
  silently break the "leave it open" behavior. Repeated test runs still
  accumulate open Chromium windows (no more `chromedriver.exe`, since
  Playwright talks to the browser directly rather than through a separate
  driver-server process — but `chrome.exe` process buildup is still possible;
  check for it the same way as before if automation runs start behaving
  erratically).
- **Repeated exports in the same session get de-duplicated filenames** (`Sent
  Cases.csv`, `Sent Cases (1).csv`, `Sent Cases (2).csv`, ...) via
  `export_and_download()`'s own counter loop against `dest_path.exists()` in
  `MIS_EXPORT_DIR` — this replaced relying on Chrome's own auto-rename
  behavior (which was specific to the old polling-based download flow); it's
  been verified directly (three exports in one session against a local mock
  server, distinct filenames, all three files present and readable).
- **This rewrite has since been verified against the live MIS site, not just
  the mock.** The mock (see "Testing the Playwright rewrite" below) is what
  the *initial* Selenium → Playwright rewrite was validated against, before
  any live-site access was exercised in this project. Since then, the full
  chain — login, Host/Database/Data Time Slab/Query selection, ARS field
  fill, and Export/download — has been run against the real site repeatedly
  and confirmed working end-to-end for all three queries (Advance Tracker,
  Case History, Sent Cases) in the same session, including finding and fixing
  two real site-specific issues that the mock didn't (and structurally
  couldn't) surface: the `channel="chrome"` network-blocking issue, and the
  Data-Time-Slab-reselection-resets-Query-dropdown behavior. Don't assume
  "only verified against a mock" still applies when troubleshooting.
- **`Report Status` in the Sent Cases export is lowercase (`"sent"`), not
  `"Sent"`.** Verified against a real export (`Sent Cases.csv.zip`, ~4700
  rows: `sent` × 4694, plus small counts of `Not Sent` / `New` / `Sent for
  Rework` / etc.). The initial implementation of `add_l2_report_sent_sheet()`
  did an exact case-sensitive match against `"Sent"` and silently matched
  zero rows every time — the `L2 Report Sent` sheet was created but always
  empty, no error raised. Both the `Report Status` and `Report Type` filters
  now compare `.strip().lower()` against `REPORT_SENT_STATUS.lower()` /
  `REPORT_SENT_TYPE.lower()` specifically because of this — don't revert to
  exact `==` matching on these two columns.
  Everything else in that function was already correct against the real file:
  header names (`case_ars_no`, `report_sent_on`, `Report Type`,
  `report_severity`, `Report Status`) match exactly, `Report Type ==
  "Additional"` matches exactly (2295 of 4710 rows), `report_sent_on` is
  `YYYY-MM-DD HH:MM:SS` (the first parse format already tried), and
  `case_ars_no` values look like `3055-016843` (matches `Pre_Offer_checkwise`
  column C). Filtering to `Report Status == "sent"` AND `Report Type ==
  "Additional"` yields 2279 rows on this export.

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

## Testing the Playwright rewrite

Live MIS site access turned out to be available from this dev environment
after all (credentials are hardcoded in `mphasis.py`, and the site was
reachable directly) — most debugging in this project ended up happening
against the real site rather than a mock, and is the stronger source of
truth where the two disagree. Early on, though, before that was established,
the initial Selenium → Playwright rewrite was validated against a small
local `http.server`-based mock replicating the MIS site's key behaviors: a
login form, Host/Database/Data-Time-Slab/Query dropdowns that rebuild
asynchronously via `onchange` JS (~400ms delay, matching the real site), a
`table#date` ARS-field container that rebuilds specifically on Query change,
and an `/export` endpoint that sleeps a few seconds before responding with
`Content-Disposition: attachment` (simulating slow server-side report
generation — the exact scenario that used to hang Selenium).

The mock is still useful as a first-pass sanity check when iterating on
automation *mechanics* in isolation (faster feedback loop, no risk of
hammering the real site) — but it's not a substitute for live verification,
and it structurally can't surface real site quirks it wasn't built to
simulate (it missed both the `channel="chrome"` issue and the
Query-dropdown-reset issue, which only showed up live). If MIS automation
needs debugging again, prefer testing against the live site directly first;
fall back to a mock like this only if live access isn't available. Two
non-obvious things that cost real time building the original mock harness,
worth not re-discovering the hard way:
- **XPath position predicates apply per-parent-context, not across the
  flattened result set from the previous step.** For `/a/b/c[2]`, if step
  `/b` yields matches under multiple different parents, `c[2]` is evaluated
  *separately within each parent's own children* — it does not mean "the 2nd
  match overall." Getting this wrong when hand-building test HTML produced a
  DOM that looked structurally similar to the intended target but resolved
  to zero or wrong matches for the real site's verified XPath
  (`/html/body/div/div/div[2]/form/input`, which needs **three** levels of
  div nesting to resolve uniquely, not two — easy to undercount).
- **`page.goto()` to a URL that triggers a download raises `"Download is
  starting"` as an error**, even inside `page.expect_download()` — this is
  why the code uses `page.click()` on the form's submit button instead.
  Confirmed directly: swapping `page.click(selector)` for
  `page.goto(export_url)` inside the same `expect_download()` block
  reproduces this error immediately. Don't refactor `export_and_download()`
  to navigate directly to an export URL even if it looks simpler.
