#!/usr/bin/env python3
"""
Fill in the "Offered Fall 2026" / "Offered Winter 2027" columns of the CIR
approved course list using the UChicago public class search.

A course counts as offered in a term if the primary course OR any of its
cross-listed courses (CL Subject / CL Course Number pairs) has a result row.

Setup (one time, in Terminal):
    pip3 install openpyxl playwright
    python3 -m playwright install chromium

Run:
    python3 check_courses.py --limit 5   # test on the first 5 rows
    python3 check_courses.py             # everything

Your original file is never changed; results go to "... - checked.xlsx".
"""
import argparse
import os
import re
import sys
import time

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

# ----------------------------- CONFIG ---------------------------------------
EXCEL_PATH = "/Users/winterb/Downloads/CIR Master Approved Course List - annotated.xlsx"
SHEET_NAME = None  # None = first sheet

URL = ("https://coursesearch92.ais.uchicago.edu/psc/prd92guest/EMPLOYEE/HRMS/c/"
       "UC_STUDENT_RECORDS_FL.UC_CLASS_SEARCH_FL.GBL")

# Spreadsheet header -> value in the site's Term dropdown
#   2268 = Autumn 2026, 2272 = Winter 2027
TERM_COLUMNS = {
    "Offered Winter 2027": "2272",
}

HEADER_SUBJECT = "Subject Code"
HEADER_NUMBER = "Number"
HEADER_CL_SUBJECT = "CL Subject"
HEADER_CL_NUMBER = "CL Course Number"

# Selectors taken from the page's HTML
SEL_TERM = "#UC_CLSRCH_WRK2_STRM"
SEL_KEYWORD = "#UC_CLSRCH_WRK2_PTUN_KEYWORD"
SEL_SEARCH_BTN = "#UC_CLSRCH_WRK_SSR_PB_SEARCH"
SEL_RESULT_ROWS = "tr[id^='DESCR100$0_row_']"
SEL_MESSAGE = "#DERIVED_CLSMSG_ERROR_TEXT"

DEBUG = False
HEADLESS = True
PAUSE_SECONDS = 1.0
NETWORK_RETRY_SEC = 20        # wait this long between tries when offline
NETWORK_WAIT_MAX_SEC = 1800   # give up (and save) after this long offline
SAVE_EVERY = 25  # save progress to the checked file every N rows
RESULT_TIMEOUT_MS = 10000
# -----------------------------------------------------------------------------


def clean_num(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def norm(v):
    """Lowercase and collapse all whitespace (incl. line breaks, nbsp)."""
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip().lower()


def find_header(wb):
    """Search every sheet for the header row; return (worksheet, row_number)."""
    targets = [wb[SHEET_NAME]] if SHEET_NAME else wb.worksheets
    for ws in targets:
        for row in ws.iter_rows(min_row=1, max_row=60):
            for c in row:
                if c.value is not None and norm(c.value) == norm(HEADER_SUBJECT):
                    return ws, c.row
    print(f'Could not find a header cell matching "{HEADER_SUBJECT}".')
    print("Sheets in this workbook:", wb.sheetnames)
    for ws in targets:
        print(f"\nFirst rows of sheet '{ws.title}':")
        for row in ws.iter_rows(min_row=1, max_row=6, values_only=True):
            print([v for v in row if v is not None][:8])
    sys.exit(1)


def is_network_error(msg):
    return any(k in msg for k in (
        "ERR_INTERNET_DISCONNECTED", "ERR_NAME_NOT_RESOLVED",
        "ERR_NETWORK_CHANGED", "ERR_CONNECTION", "ERR_PROXY",
        "ERR_ADDRESS_UNREACHABLE", "ERR_TIMED_OUT"))


def settle(page):
    page.wait_for_load_state("networkidle")
    time.sleep(0.7)


def search(page, subject, number, term_value):
    """Return 'Yes', 'No', or 'Unclear' for one course in one term."""
    page.goto(URL, wait_until="networkidle")

    # Changing the term triggers a page update, so wait for it to settle
    page.select_option(SEL_TERM, value=term_value)
    settle(page)

    page.fill(SEL_KEYWORD, f"{subject} {number}")
    page.click(SEL_SEARCH_BTN)

    # Let the search run, then watch the page's visible text for the course code
    settle(page)
    pattern = re.compile(rf"\b{re.escape(subject)}\s*{re.escape(number)}\b", re.I)
    deadline = time.time() + RESULT_TIMEOUT_MS / 1000
    found = False
    body = ""
    quiet_polls = 0
    while time.time() < deadline:
        body = page.inner_text("body")
        if pattern.search(body):
            found = True
            break
        # If the site has clearly finished (message or result count shown)
        # and the course still isn't there, stop waiting early.
        if site_finished(page):
            quiet_polls += 1
            if quiet_polls >= 3:
                break
        time.sleep(0.5)

    if DEBUG:
        base = f"search_debug_{subject}_{number}"
        with open(base + ".txt", "w") as f:
            f.write(body)
        try:
            page.screenshot(path=base + ".png", full_page=True)
        except Exception:
            pass
        print(f"   DEBUG {subject} {number}: found={found}; "
              f"saved {base}.txt and {base}.png in the current folder")

    return "Yes" if found else "No"


def site_finished(page):
    """True once the site shows a message or a results count."""
    for sel in (SEL_MESSAGE, "#UC_RSLT_NAV_WRK_PTPG_ROWS_GRID"):
        try:
            t = page.locator(sel).first.inner_text(timeout=500)
            if t.replace("\xa0", " ").strip():
                return True
        except Exception:
            pass
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--debug", action="store_true",
                    help="print what the site returns for each search")
    ap.add_argument("--retry-unclear", action="store_true",
                    help="re-check only rows marked Unclear in the existing "
                         "'- checked.xlsx' file")
    ap.add_argument("--start-at", type=int, default=1,
                    help="continue from this course number (1-based) using "
                         "the existing '- checked.xlsx' file")
    args = ap.parse_args()
    global DEBUG
    DEBUG = args.debug

    out = EXCEL_PATH.replace(".xlsx", " - checked.xlsx")
    resuming = args.retry_unclear or args.start_at > 1
    if resuming:
        if not os.path.isfile(out):
            sys.exit(f"Can't resume: no checked file found at {out}")
        print(f"Resuming from existing results file: {out}")
    wb = load_workbook(out if resuming else EXCEL_PATH)
    ws, hdr = find_header(wb)
    print(f"Using sheet '{ws.title}', header row {hdr}")

    normcols = {}
    for c in ws[hdr]:
        if c.value is not None:
            normcols.setdefault(norm(c.value), []).append(c.column)

    # Expose columns under the names used in CONFIG (matched loosely)
    print("Headers found in row:", list(normcols))

    cols = {}
    for name in [HEADER_SUBJECT, HEADER_NUMBER, HEADER_CL_SUBJECT,
                 HEADER_CL_NUMBER, *TERM_COLUMNS]:
        if norm(name) in normcols:
            cols[name] = normcols[norm(name)]

    # Term columns: fall back to a loose match, else create the column
    for name in TERM_COLUMNS:
        if name in cols:
            continue
        words = [w for w in norm(name).split() if w != "offered"]
        loose = [k for k in normcols if all(w in k for w in words)]
        if loose:
            cols[name] = normcols[loose[0]]
            print(f'Using existing column "{loose[0]}" for "{name}"')
        else:
            new_col = ws.max_column + 1
            ws.cell(hdr, new_col).value = name
            cols[name] = [new_col]
            print(f'Column "{name}" not found; added it as a new column.')

    for needed in [HEADER_SUBJECT, HEADER_NUMBER, *TERM_COLUMNS]:
        if needed not in cols:
            sys.exit(f'Column "{needed}" not found. Headers seen: {list(cols)}')

    subj_col, num_col = cols[HEADER_SUBJECT][0], cols[HEADER_NUMBER][0]
    cl_pairs = list(zip(cols.get(HEADER_CL_SUBJECT, []), cols.get(HEADER_CL_NUMBER, [])))

    jobs = []
    for r in range(hdr + 1, ws.max_row + 1):
        s, n = ws.cell(r, subj_col).value, ws.cell(r, num_col).value
        if not s or n is None:
            continue
        courses = [(str(s).strip().upper(), clean_num(n))]
        for sc, nc in cl_pairs:
            cs, cn = ws.cell(r, sc).value, ws.cell(r, nc).value
            if cs and cn is not None:
                courses.append((str(cs).strip().upper(), clean_num(cn)))
        jobs.append((r, courses))
    if args.start_at > 1:
        jobs = jobs[args.start_at - 1:]
    if args.retry_unclear:
        jobs = [(r, c) for r, c in jobs
                if any(ws.cell(r, cols[h][0]).value == "Unclear" for h in TERM_COLUMNS)]
        print(f"{len(jobs)} row(s) marked Unclear will be re-checked.")
    if args.limit:
        jobs = jobs[: args.limit]

    cache = {}

    current = {"i": 0, "row": None}

    def abort_for_network():
        """Save progress and stop cleanly if the connection stays down."""
        if current["row"] is not None:
            for h in TERM_COLUMNS:
                ws.cell(current["row"], cols[h][0]).value = "Unclear"
        wb.save(out)
        print("\nThe internet connection has been down too long. Progress was saved.")
        again = ("--retry-unclear" if args.retry_unclear
                 else f"--start-at {current['i']}")
        print(f"Once you're back online, continue with:\n"
              f"    python3 check_courses.py {again}")
        sys.exit(1)

    def lookup(page, course, term_value):
        key = (course, term_value)
        if key in cache:
            return cache[key]
        waited, tries = 0, 0
        while True:
            try:
                cache[key] = search(page, course[0], course[1], term_value)
                break
            except Exception as e:
                msg = str(e)
                if is_network_error(msg):
                    if waited >= NETWORK_WAIT_MAX_SEC:
                        abort_for_network()
                    print(f"   no internet connection; waiting {NETWORK_RETRY_SEC}s "
                          f"and retrying ({course[0]} {course[1]})...")
                    time.sleep(NETWORK_RETRY_SEC)
                    waited += NETWORK_RETRY_SEC
                    continue
                tries += 1
                if tries < 2:  # one quiet retry for ordinary hiccups
                    time.sleep(2)
                    continue
                print(f"   error on {course}: {msg[:120]}")
                cache[key] = "Unclear"
                break
        time.sleep(PAUSE_SECONDS)
        return cache[key]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_context().new_page()
        page.set_default_timeout(20000)

        for i, (row, courses) in enumerate(jobs, 1):
            current["i"] = i + max(args.start_at - 1, 0) if not args.retry_unclear else i
            current["row"] = row
            label = " / ".join(f"{s} {n}" for s, n in courses)
            for header, term_value in TERM_COLUMNS.items():
                result = "No"
                for course in courses:
                    res = lookup(page, course, term_value)
                    if res == "Yes":
                        result = "Yes"
                        break
                    if res == "Unclear":
                        result = "Unclear"
                # Spreadsheet: "Y" if offered, blank if not, "Unclear" on errors
                cell_value = {"Yes": "Y", "No": None}.get(result, result)
                ws.cell(row, cols[header][0]).value = cell_value
                print(f"[{i}/{len(jobs)}] {label} | {header}: {cell_value or '(blank)'}")

            if i % SAVE_EVERY == 0:
                wb.save(out)
                print(f"   (progress saved after {i} rows)")

        browser.close()

    wb.save(out)
    print(f"\nSaved: {out}")


if __name__ == "__main__":
    main()
