#!/usr/bin/env python3
"""ServiceNow open-incident scraper and Andersons-branded HTML/Excel reporter.

Scrapes the SN - Corporate Solutions incident list from ServiceNow (Playwright,
interactive login), filters out Resolved rows, sorts by priority, and writes:

  - Excel (.xlsx)
  - Self-contained HTML report (Outlook-paste friendly)

Usage:
  pip install -r requirements-snow.txt
  playwright install chromium
  python scripts/servicenow_incident_report.py

Optional environment overrides:
  SNOW_OUTPUT_DIR   Directory for Excel/HTML output (default: /home/ANDE/SNOW)
  SNOW_LOGO_PATH    Path to Andersons logo image for the HTML header
"""

from __future__ import annotations

import base64
import html as html_lib
import os
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from playwright.sync_api import Frame, Page, TimeoutError, sync_playwright

OUTPUT_DIR = Path(os.environ.get("SNOW_OUTPUT_DIR", "/home/ANDE/SNOW"))
FALLBACK_OUTPUT_DIR = Path.home() / "SNOW"
URL = (
    "https://andersonsinc.service-now.com/now/nav/ui/classic/params/target/"
    "incident_list.do%3Fsysparm_first_row%3D1%26sysparm_query%3Dactive%253dtrue%255e"
    "GOTOassignment_group.name%253e%253dSN%2B-%2BCorporate%2BSolutions%26"
    "sysparm_query_encoded%3Dactive%253dtrue%255eGOTOassignment_group.name%253e%253d"
    "SN%2B-%2BCorporate%2BSolutions%26sysparm_view%3D"
)
LOGIN_TIMEOUT_MS = 300_000
PAGE_LOAD_WAIT_MS = 10_000

# --- Andersons / TCS branding ---
ANDE_BLUE = "#002D5B"
ANDE_GOLD = "#FFB432"
ANDE_CREAM = "#FAF8F4"
# Optional: point this at a local Andersons logo (png/jpg/svg). If the file
# exists it is base64-embedded into the HTML; otherwise a styled text
# wordmark is rendered so the report never breaks.
LOGO_PATH = Path(
    os.environ.get("SNOW_LOGO_PATH", "/home/harishgupta/ANDE/andersons_logo.png")
)

PRIORITY_COLORS = {
    1: ("#C00000", "#FDECEC"),  # Critical  - red
    2: ("#C55A11", "#FDF3E7"),  # High      - orange
    3: ("#BF8F00", "#FFF8E1"),  # Moderate  - amber
    4: ("#548235", "#EFF6EA"),  # Low       - green
    5: ("#2E75B6", "#EAF1F8"),  # Planning  - blue
}

PRIORITY_RANK = {
    "1 - critical": 1,
    "critical": 1,
    "1": 1,
    "2 - high": 2,
    "high": 2,
    "2": 2,
    "3 - moderate": 3,
    "3 - medium": 3,
    "moderate": 3,
    "medium": 3,
    "3": 3,
    "4 - low": 4,
    "low": 4,
    "4": 4,
    "5 - planning": 5,
    "planning": 5,
    "5": 5,
}


def normalize_text(value: str) -> str:
    return " ".join(value.replace("\xa0", " ").split()).strip()


def find_list_frame(page: Page) -> Frame:
    page.wait_for_timeout(PAGE_LOAD_WAIT_MS)

    selectors = [
        "iframe[name='gsft_main']",
        "iframe#gsft_main",
        "frame[name='gsft_main']",
    ]

    for selector in selectors:
        try:
            locator = page.locator(selector)
            if locator.count() > 0:
                return page.frame(name="gsft_main") or page.frames[0]
        except Exception:
            continue

    return page.main_frame


def wait_for_successful_login(page: Page) -> None:
    print("Waiting for successful login...")

    success_indicators = [
        "iframe[name='gsft_main']",
        "frame[name='gsft_main']",
        "table.list_table",
        "table.data_list_table",
        "#incident_table",
        "input[id*='sysparm_search']",
    ]

    page.wait_for_load_state("domcontentloaded")
    try:
        page.wait_for_selector(
            ", ".join(success_indicators),
            timeout=LOGIN_TIMEOUT_MS,
            state="visible",
        )
        return
    except TimeoutError as exc:
        raise RuntimeError(
            "Login was not detected within the timeout. Complete authentication in the browser and ensure the incident list loads."
        ) from exc


def get_table_locator(frame: Frame):
    for selector in [
        "table.list_table",
        "table.data_list_table",
        "table.list2_table",
        "table#incident_table",
        "table[role='presentation']",
        "table",
    ]:
        locator = frame.locator(selector)
        if locator.count() > 0:
            return locator.first
    raise RuntimeError("Could not locate the incident list table.")


def extract_headers(table) -> list[str]:
    # ServiceNow thead contains BOTH a label row and a search/filter row whose
    # cells all read "Search". Score candidate rows by the number of UNIQUE
    # non-empty labels: the search row scores 1 ("Search" repeated) and can
    # never beat the real label row. Empty headers (checkbox/preview columns)
    # become Column_N placeholders to preserve positional alignment.
    header_rows = table.locator("thead tr")
    best_names: list[str] = []
    best_score = -1

    for r in range(header_rows.count()):
        ths = header_rows.nth(r).locator("th")
        names: list[str] = []
        labels: list[str] = []
        for i in range(ths.count()):
            text = normalize_text(ths.nth(i).inner_text())
            if text:
                labels.append(text)
                names.append(text)
            else:
                names.append(f"Column_{i + 1}")
        score = len(set(label.lower() for label in labels))
        if score > best_score and names:
            best_score = score
            best_names = names

    if best_names and best_score > 0:
        # Dedupe repeated header names so dict keys never collide and
        # silently swallow column data.
        seen: dict[str, int] = {}
        deduped: list[str] = []
        for name in best_names:
            key = name.lower()
            if key in seen:
                seen[key] += 1
                deduped.append(f"{name}_{seen[key]}")
            else:
                seen[key] = 1
                deduped.append(name)
        return deduped

    first_row_cells = table.locator("tbody tr").first.locator("td")
    return [f"Column_{i + 1}" for i in range(first_row_cells.count())]


def read_cell_text(cell) -> str:
    """inner_text first (rendered text); fall back to text_content for cells
    ServiceNow renders as truncated/hidden (a known cause of blank
    Short description values)."""
    text = normalize_text(cell.inner_text())
    if text:
        return text
    try:
        return normalize_text(cell.text_content() or "")
    except Exception:
        return ""


def extract_rows_from_current_page(frame: Frame) -> list[dict[str, str]]:
    table = get_table_locator(frame)
    headers = extract_headers(table)
    body_rows = table.locator("tbody tr")
    if body_rows.count() == 0:
        body_rows = table.locator("tr")

    records: list[dict[str, str]] = []
    for i in range(body_rows.count()):
        row = body_rows.nth(i)
        cells = row.locator("td")
        cell_count = cells.count()
        if cell_count == 0:
            continue

        values = [read_cell_text(cells.nth(j)) for j in range(cell_count)]
        if not any(values):
            continue

        if len(headers) < cell_count:
            headers = headers + [f"Column_{idx + 1}" for idx in range(len(headers), cell_count)]

        record = {headers[j]: values[j] for j in range(cell_count)}
        records.append(record)

    return records


def get_row_identity(record: dict[str, str]) -> str:
    for key in ["Number", "Task number", "Incident", "Sys ID", "Column_1"]:
        value = record.get(key, "")
        if value:
            return value
    return "|".join(record.values())


def click_next_page(frame: Frame) -> bool:
    next_selectors = [
        "button[name='vcr_next']",
        "a[aria-label*='Next page']",
        "button[aria-label*='Next page']",
        "a[title*='Next page']",
        "a.list2_next",
        "a.icon-vcr-right",
        "a[onclick*='next']",
    ]

    rows_locator = frame.locator("tbody tr")
    if rows_locator.count() == 0:
        rows_locator = frame.locator("tr")
    current_marker = normalize_text(rows_locator.first.inner_text())

    for selector in next_selectors:
        locator = frame.locator(selector)
        count = locator.count()
        if count == 0:
            continue

        for i in range(count):
            candidate = locator.nth(i)

            try:
                if not candidate.is_visible():
                    continue
            except Exception:
                continue

            classes = (candidate.get_attribute("class") or "").lower()
            aria_disabled = (candidate.get_attribute("aria-disabled") or "").lower()
            disabled_attr = candidate.get_attribute("disabled")
            if "disabled" in classes or aria_disabled == "true" or disabled_attr is not None:
                continue

            try:
                candidate.click(timeout=5_000)
            except TimeoutError:
                continue

            try:
                frame.page.wait_for_timeout(2_500)
                rows_locator = frame.locator("tbody tr")
                if rows_locator.count() == 0:
                    rows_locator = frame.locator("tr")
                rows_locator.first.wait_for(state="visible", timeout=15_000)
                new_marker = normalize_text(rows_locator.first.inner_text())
                return new_marker != current_marker
            except TimeoutError:
                return False

    return False


def scrape_all_pages(frame: Frame) -> list[dict[str, str]]:
    all_records: list[dict[str, str]] = []
    seen_ids: set[str] = set()
    stagnant_pages = 0
    page_number = 1

    while True:
        current_records = extract_rows_from_current_page(frame)

        if page_number == 1 and current_records:
            print(f"Detected columns: {list(current_records[0].keys())}")
            print(f"First raw row: {current_records[0]}")

        added_this_page = 0

        for record in current_records:
            row_id = get_row_identity(record)
            if row_id in seen_ids:
                continue
            seen_ids.add(row_id)
            all_records.append(record)
            added_this_page += 1

        print(f"Page {page_number}: {added_this_page} new rows (total {len(all_records)})")

        if added_this_page == 0:
            stagnant_pages += 1
        else:
            stagnant_pages = 0

        if stagnant_pages >= 2:
            print("Two consecutive pages with no new rows - stopping pagination.")
            break

        try:
            advanced = click_next_page(frame)
        except Exception as exc:
            print(f"Pagination stopped due to error ({exc}). Keeping {len(all_records)} rows scraped so far.")
            break

        if not advanced:
            print("No further visible/enabled Next button - assuming last page.")
            break

        page_number += 1

    return all_records


def find_column(columns: Iterable[str], *candidates: str) -> str | None:
    normalized = {column.lower(): column for column in columns}
    for candidate in candidates:
        if candidate.lower() in normalized:
            return normalized[candidate.lower()]
    return None


def drop_placeholder_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Remove decoration columns (checkbox/preview) from the output entirely:
    any column that is all-empty, and any Column_N placeholder that carries
    no meaningful data."""
    if df.empty:
        return df
    keep = []
    for col in df.columns:
        values = df[col].astype(str).str.strip()
        all_empty = values.eq("").all()
        if all_empty:
            continue
        keep.append(col)
    return df[keep]


def warn_on_missing_columns(df: pd.DataFrame) -> None:
    """Surface data-quality problems loudly instead of exporting silently."""
    checks = {
        "Assigned to": (
            "'Assigned to' column NOT found. The scraper can only capture columns "
            "shown in the ServiceNow list view. Add it via the gear icon -> "
            "Personalize List -> add 'Assigned to', then rerun."
        ),
        "Short description": (
            "'Short description' column NOT found in the list view. Add it via "
            "the gear icon -> Personalize List if needed."
        ),
    }
    for name, message in checks.items():
        col = find_column(df.columns, name)
        if col is None:
            print(f"WARNING: {message}")
        else:
            empty_count = df[col].astype(str).str.strip().eq("").sum()
            if len(df) > 0 and empty_count == len(df):
                print(
                    f"WARNING: '{col}' column exists but every value is empty - "
                    f"the cell text is not being captured. Check the 'First raw row' "
                    f"diagnostic above and share it if this persists."
                )
            elif empty_count > 0:
                print(f"Note: '{col}' is empty on {empty_count} of {len(df)} rows.")


def filter_and_sort_records(records: list[dict[str, str]]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df.columns = [normalize_text(str(column)) for column in df.columns]

    assignment_col = find_column(df.columns, "Assignment group", "Assigned to group")
    state_col = find_column(df.columns, "State", "Status")
    priority_col = find_column(df.columns, "Priority")

    if assignment_col is not None:
        df = df[df[assignment_col].astype(str).str.strip().eq("SN - Corporate Solutions")]

    if state_col is not None:
        df = df[~df[state_col].astype(str).str.strip().str.upper().eq("RESOLVED")]

    if priority_col is not None:
        df["__priority_rank"] = (
            df[priority_col]
            .astype(str)
            .str.strip()
            .str.lower()
            .map(PRIORITY_RANK)
            .fillna(999)
        )
        df = df.sort_values(by=["__priority_rank", priority_col], ascending=[True, True])
        df = df.drop(columns=["__priority_rank"])

    df = drop_placeholder_columns(df)
    df = df.reset_index(drop=True)
    return df


def priority_rank_of(value: str) -> int | None:
    return PRIORITY_RANK.get(str(value).strip().lower())


def build_logo_html() -> str:
    """Base64-embed the Andersons logo if available; else a styled wordmark."""
    if LOGO_PATH.is_file():
        try:
            data = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
            suffix = LOGO_PATH.suffix.lower().lstrip(".")
            mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                    "gif": "image/gif", "svg": "image/svg+xml"}.get(suffix, "image/png")
            return (
                f'<img src="data:{mime};base64,{data}" alt="The Andersons" '
                f'style="height:48px;vertical-align:middle;" />'
            )
        except OSError:
            pass
    return (
        f'<span style="font-family:Georgia,serif;font-size:26px;font-weight:bold;'
        f'color:{ANDE_BLUE};vertical-align:middle;">The Andersons'
        f'<span style="color:{ANDE_GOLD};">&#9650;</span></span>'
    )


def write_html_report(df: pd.DataFrame, output_file: Path) -> None:
    """Self-contained HTML report in Andersons branding. Pure inline CSS,
    no JavaScript/CDN, so it can be pasted into Outlook (Cmd+A, Cmd+C)."""
    generated = datetime.now().strftime("%B %d, %Y %I:%M %p")

    # Hide placeholder columns (e.g. the ServiceNow checkbox column) that are
    # entirely empty - display only; the Excel export keeps all columns.
    if not df.empty:
        display_cols = [
            col for col in df.columns
            if not (str(col).startswith("Column_")
                    and df[col].astype(str).str.strip().eq("").all())
        ]
        df = df[display_cols]

    priority_col = find_column(df.columns, "Priority") if not df.empty else None

    # Summary counts by priority
    counts: dict[int, int] = {}
    if priority_col is not None:
        for value in df[priority_col]:
            rank = priority_rank_of(value)
            if rank is not None:
                counts[rank] = counts.get(rank, 0) + 1

    summary_cards = ""
    labels = {1: "Critical", 2: "High", 3: "Medium", 4: "Low", 5: "Planning"}
    for rank in [1, 2, 3, 4, 5]:
        fg, bg = PRIORITY_COLORS[rank]
        summary_cards += (
            f'<td style="background:{bg};border:1px solid {fg};border-radius:6px;'
            f'padding:10px 18px;text-align:center;">'
            f'<div style="font-size:26px;font-weight:bold;color:{fg};">{counts.get(rank, 0)}</div>'
            f'<div style="font-size:11px;color:{ANDE_BLUE};text-transform:uppercase;'
            f'letter-spacing:1px;">{labels[rank]}</div></td>'
            f'<td style="width:10px;"></td>'
        )

    # Table
    if df.empty:
        table_html = (
            f'<p style="color:{ANDE_BLUE};font-style:italic;">'
            f'No open incidents matched the filters (SN - Corporate Solutions, not Resolved).</p>'
        )
    else:
        header_cells = "".join(
            f'<th style="background:{ANDE_BLUE};color:#FFFFFF;padding:8px 10px;'
            f'text-align:left;font-size:12px;border-bottom:3px solid {ANDE_GOLD};'
            f'white-space:nowrap;">{html_lib.escape(str(col))}</th>'
            for col in df.columns
        )
        body_rows = ""
        for i, (_, row) in enumerate(df.iterrows()):
            rank = priority_rank_of(row[priority_col]) if priority_col is not None else None
            zebra = "#FFFFFF" if i % 2 == 0 else ANDE_CREAM
            cells = ""
            for col in df.columns:
                value = html_lib.escape(str(row[col]))
                style = (
                    'padding:7px 10px;font-size:12px;color:#333333;'
                    'border-bottom:1px solid #E3DED6;vertical-align:top;'
                )
                if col == priority_col and rank in PRIORITY_COLORS:
                    fg, bg = PRIORITY_COLORS[rank]
                    style += (
                        f'color:{fg};background:{bg};font-weight:bold;white-space:nowrap;'
                    )
                cells += f'<td style="{style}">{value}</td>'
            body_rows += f'<tr style="background:{zebra};">{cells}</tr>'
        table_html = (
            f'<table cellpadding="0" cellspacing="0" '
            f'style="border-collapse:collapse;width:100%;border:1px solid {ANDE_BLUE};">'
            f'<thead><tr>{header_cells}</tr></thead><tbody>{body_rows}</tbody></table>'
        )

    html_doc = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<title>ServiceNow Open Incidents - SN - Corporate Solutions</title>
</head>
<body style="margin:0;padding:0;background:{ANDE_CREAM};font-family:Calibri,Arial,sans-serif;">
<table cellpadding="0" cellspacing="0" style="width:100%;max-width:1100px;margin:0 auto;background:{ANDE_CREAM};">
<tr><td style="padding:24px 28px 0 28px;">
  <table cellpadding="0" cellspacing="0" style="width:100%;">
  <tr>
    <td style="vertical-align:middle;">{build_logo_html()}</td>
    <td style="text-align:right;vertical-align:middle;font-family:Calibri,Arial,sans-serif;
        font-size:12px;color:{ANDE_BLUE};">Generated: {generated}</td>
  </tr>
  </table>
  <div style="border-bottom:3px solid {ANDE_GOLD};margin:14px 0 0 0;"></div>
</td></tr>
<tr><td style="padding:20px 28px 6px 28px;">
  <h1 style="margin:0;font-size:22px;color:{ANDE_BLUE};font-family:Calibri,Arial,sans-serif;">
    ServiceNow Open Incident Report</h1>
  <p style="margin:4px 0 0 0;font-size:13px;color:#555555;">
    Assignment Group: <b>SN - Corporate Solutions</b> &nbsp;|&nbsp;
    Excluding Resolved &nbsp;|&nbsp; Sorted by Priority &nbsp;|&nbsp;
    Total open incidents: <b>{len(df)}</b></p>
</td></tr>
<tr><td style="padding:14px 28px 6px 28px;">
  <table cellpadding="0" cellspacing="0"><tr>{summary_cards}</tr></table>
</td></tr>
<tr><td style="padding:14px 28px 20px 28px;">{table_html}</td></tr>
<tr><td style="padding:0 28px 26px 28px;">
  <div style="border-top:2px solid {ANDE_GOLD};padding-top:10px;font-size:11px;
      color:{ANDE_BLUE};font-family:Calibri,Arial,sans-serif;">
    <b>Harish Gupta</b> | Enterprise Solutions Unit &mdash; Consulting Practice |
    gupta.h@tcs.com<br/>
    <span style="color:#777777;">TCS / The Andersons Confidential &mdash;
    Source: ServiceNow incident list extract</span>
  </div>
</td></tr>
</table>
</body>
</html>"""

    output_file.write_text(html_doc, encoding="utf-8")


def print_preview(df: pd.DataFrame) -> None:
    if df.empty:
        print("No records matched the requested filters.")
        return

    preview = df.head(20)
    print("\nFiltered incident preview (first 20 rows):")
    print(preview.to_string(index=False))


def main() -> int:
    try:
        try:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_base = OUTPUT_DIR
        except OSError:
            FALLBACK_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_base = FALLBACK_OUTPUT_DIR

        output_file = output_base / f"servicenow_incidents_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        with sync_playwright() as p:
            print("Launching Chromium...")
            browser = p.chromium.launch(headless=False, args=["--start-maximized"])
            context = browser.new_context(viewport=None)
            page = context.new_page()

            print("Opening ServiceNow incident list...")
            page.goto(URL, wait_until="domcontentloaded")

            print("Please complete authentication in the browser window.")
            wait_for_successful_login(page)

            frame = find_list_frame(page)
            print("List located. Scraping all pages...")
            records = scrape_all_pages(frame)

            if not records:
                print("No rows were extracted. The page structure may have changed, or the list did not load.")
                browser.close()
                return 1

            print(f"Total raw records scraped: {len(records)}")
            filtered_df = filter_and_sort_records(records)
            warn_on_missing_columns(filtered_df)
            try:
                filtered_df.to_excel(output_file, index=False)
            except ImportError as exc:
                raise RuntimeError(
                    "Excel export requires an engine such as openpyxl. Install it with: pip install openpyxl"
                ) from exc

            print_preview(filtered_df)
            print(f"\nFiltered Excel file saved to: {output_file.resolve()}")

            html_file = output_file.with_suffix(".html")
            try:
                write_html_report(filtered_df, html_file)
                print(f"Andersons-branded HTML report saved to: {html_file.resolve()}")
                if not LOGO_PATH.is_file():
                    print(f"Note: logo file not found at {LOGO_PATH} - used styled text wordmark instead.")
            except Exception as exc:
                print(f"HTML report generation failed ({exc}). Excel output is unaffected.")

            print(f"Filtered rows exported: {len(filtered_df)}")

            browser.close()
            return 0

    except Exception as exc:
        print(f"Extraction failed: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
