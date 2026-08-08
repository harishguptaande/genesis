#!/usr/bin/env python3
"""ServiceNow open-incident scraper and Project Genesis branded reporter.

Scrapes the SN - Corporate Solutions incident list from ServiceNow (Playwright,
interactive login), filters out Resolved rows, sorts by priority (On Hold last),
and writes:

  - Excel (.xlsx)
  - Self-contained HTML report (Outlook-paste friendly)
  - 16:9 executive dashboard HTML
  - PowerPoint summary slide (.pptx, optional via python-pptx)

Usage:
  pip install -r requirements-snow.txt
  playwright install chromium
  python scripts/servicenow_incident_report.py
"""

from __future__ import annotations

import html as html_lib
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from playwright.sync_api import Frame, Page, TimeoutError, sync_playwright

OUTPUT_DIR = Path("/home/ANDE/SNOW")
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

# --- Project Genesis / TCS branding ---
ANDE_BLUE = "#002D5B"
ANDE_GOLD = "#FFB432"
ANDE_CREAM = "#FAF8F4"
# Program wordmark shown in place of a logo on all report headers.
PROGRAM_NAME = "Project Genesis"

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
    """Remove ServiceNow decoration columns from the output entirely:
    every Column_N placeholder (unlabeled columns - row selection checkbox,
    preview icon) regardless of content, plus any column that is all-empty."""
    if df.empty:
        return df
    keep = []
    for col in df.columns:
        if str(col).startswith("Column_"):
            continue
        values = df[col].astype(str).str.strip()
        if values.eq("").all():
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

    sort_cols: list[str] = []

    # On Hold items go LAST regardless of priority (least importance).
    if state_col is not None:
        df["__onhold"] = (
            df[state_col].astype(str).str.strip().str.lower().eq("on hold").astype(int)
        )
        sort_cols.append("__onhold")

    if priority_col is not None:
        df["__priority_rank"] = (
            df[priority_col]
            .astype(str)
            .str.strip()
            .str.lower()
            .map(PRIORITY_RANK)
            .fillna(999)
        )
        sort_cols.append("__priority_rank")
        sort_cols.append(priority_col)

    if sort_cols:
        df = df.sort_values(by=sort_cols, ascending=True, kind="stable")
        df = df.drop(columns=[c for c in ["__onhold", "__priority_rank"] if c in df.columns])

    df = drop_placeholder_columns(df)
    df = df.reset_index(drop=True)
    return df


def priority_rank_of(value: str) -> int | None:
    return PRIORITY_RANK.get(str(value).strip().lower())


def build_wordmark_html(font_px: int = 26) -> str:
    """Project Genesis wordmark in Andersons brand colors (no external image,
    so the header always renders identically in email, browser, and print)."""
    return (
        f'<span style="font-family:Calibri,Arial,sans-serif;font-size:{font_px}px;'
        f'font-weight:bold;color:{ANDE_BLUE};letter-spacing:0.5px;'
        f'vertical-align:middle;">{PROGRAM_NAME}'
        f'<span style="color:{ANDE_GOLD};">.</span></span>'
    )


def write_html_report(df: pd.DataFrame, output_file: Path) -> None:
    """Self-contained HTML report in Project Genesis branding. Pure inline CSS,
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
        state_col_report = find_column(df.columns, "State", "Status")
        for i, (_, row) in enumerate(df.iterrows()):
            rank = priority_rank_of(row[priority_col]) if priority_col is not None else None
            is_onhold = (
                state_col_report is not None
                and str(row[state_col_report]).strip().lower() == "on hold"
            )
            zebra = "#F1EEE8" if is_onhold else ("#FFFFFF" if i % 2 == 0 else ANDE_CREAM)
            cells = ""
            for col in df.columns:
                value = html_lib.escape(str(row[col]))
                style = (
                    'padding:7px 10px;font-size:12px;color:#333333;'
                    'border-bottom:1px solid #E3DED6;vertical-align:top;'
                )
                if is_onhold:
                    style += 'color:#777777;'
                if col == priority_col and rank in PRIORITY_COLORS:
                    fg, bg = PRIORITY_COLORS[rank]
                    if is_onhold:
                        style += f'color:{fg};opacity:0.55;font-weight:bold;white-space:nowrap;'
                    else:
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
    <td style="vertical-align:middle;">{build_wordmark_html(26)}</td>
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
    <span style="color:#777777;">TCS / Project Genesis Confidential &mdash;
    Source: ServiceNow incident list extract</span>
  </div>
</td></tr>
</table>
</body>
</html>"""

    output_file.write_text(html_doc, encoding="utf-8")


def write_dashboard_html(df: pd.DataFrame, output_file: Path) -> None:
    """Executive dashboard sized 16:9 (1280x720) as a single slide.
    Pure inline CSS, no JavaScript - copy/paste or screenshot into PowerPoint."""
    generated = datetime.now().strftime("%B %d, %Y")
    priority_col = find_column(df.columns, "Priority")
    state_col = find_column(df.columns, "State", "Status")
    assigned_col = find_column(df.columns, "Assigned To", "Assigned to")
    number_col = find_column(df.columns, "Number", "Task number")
    desc_col = find_column(df.columns, "Short Description", "Short description")

    total = len(df)
    counts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}
    if priority_col is not None:
        for value in df[priority_col]:
            rank = priority_rank_of(value)
            if rank in counts:
                counts[rank] += 1
    if state_col is not None and not df.empty:
        onhold_mask = df[state_col].astype(str).str.strip().str.lower().eq("on hold")
    else:
        onhold_mask = pd.Series([False] * len(df), index=df.index, dtype=bool)
    onhold_count = int(onhold_mask.sum())

    def kpi(value, label, fg, bg):
        return (
            f'<td style="width:16%;background:{bg};border:1px solid {fg};'
            f'border-radius:8px;padding:12px 6px;text-align:center;">'
            f'<div style="font-size:34px;font-weight:bold;color:{fg};line-height:1;">{value}</div>'
            f'<div style="font-size:11px;color:{ANDE_BLUE};text-transform:uppercase;'
            f'letter-spacing:1px;margin-top:5px;">{label}</div></td><td style="width:8px;"></td>'
        )

    kpis = kpi(total, "Total Open", ANDE_BLUE, "#FFFFFF")
    for rank, label in [(1, "Critical"), (2, "High"), (3, "Medium"), (4, "Low")]:
        fg, bg = PRIORITY_COLORS[rank]
        kpis += kpi(counts[rank], label, fg, bg)
    kpis += kpi(onhold_count, "On Hold", "#777777", "#F1EEE8")

    attention_rows = ""
    if not df.empty and number_col is not None:
        active = df[~onhold_mask].head(6)
        for _, row in active.iterrows():
            rank = priority_rank_of(row[priority_col]) if priority_col is not None else None
            fg, bg = PRIORITY_COLORS.get(rank, (ANDE_BLUE, "#FFFFFF"))
            desc = str(row[desc_col]) if desc_col is not None else ""
            if len(desc) > 58:
                desc = desc[:55] + "..."
            assignee = str(row[assigned_col]) if assigned_col is not None else ""
            if not assignee.strip():
                assignee = "UNASSIGNED"
            attention_rows += (
                f'<tr>'
                f'<td style="padding:5px 8px;font-size:12px;font-weight:bold;'
                f'color:{ANDE_BLUE};white-space:nowrap;border-bottom:1px solid #E3DED6;">'
                f'{html_lib.escape(str(row[number_col]))}</td>'
                f'<td style="padding:5px 4px;border-bottom:1px solid #E3DED6;white-space:nowrap;">'
                f'<span style="background:{bg};color:{fg};font-size:10px;font-weight:bold;'
                f'padding:2px 7px;border-radius:8px;border:1px solid {fg};">'
                f'{html_lib.escape(str(row[priority_col]) if priority_col else "")}</span></td>'
                f'<td style="padding:5px 8px;font-size:11.5px;color:#333333;'
                f'border-bottom:1px solid #E3DED6;">{html_lib.escape(desc)}</td>'
                f'<td style="padding:5px 8px;font-size:11.5px;color:{ANDE_BLUE};'
                f'white-space:nowrap;border-bottom:1px solid #E3DED6;">'
                f'{html_lib.escape(assignee)}</td></tr>'
            )

    assignee_bars = ""
    if assigned_col is not None and not df.empty:
        workload = (
            df[assigned_col].astype(str).str.strip().replace("", "UNASSIGNED")
            .value_counts().head(7)
        )
        max_count = int(workload.max()) if len(workload) else 1
        for name, count in workload.items():
            width_pct = max(8, int(count / max_count * 100))
            assignee_bars += (
                f'<tr><td style="padding:3px 8px 3px 0;font-size:11.5px;color:#333333;'
                f'white-space:nowrap;width:140px;">{html_lib.escape(str(name))}</td>'
                f'<td style="padding:3px 0;"><div style="background:{ANDE_BLUE};'
                f'height:14px;width:{width_pct}%;border-radius:2px;"></div></td>'
                f'<td style="padding:3px 0 3px 8px;font-size:11.5px;font-weight:bold;'
                f'color:{ANDE_BLUE};width:24px;">{count}</td></tr>'
            )

    state_chips = ""
    if state_col is not None and not df.empty:
        for state, count in df[state_col].astype(str).str.strip().value_counts().items():
            is_hold = state.lower() == "on hold"
            chip_fg = "#777777" if is_hold else ANDE_BLUE
            chip_bg = "#F1EEE8" if is_hold else "#FFFFFF"
            state_chips += (
                f'<span style="display:inline-block;background:{chip_bg};color:{chip_fg};'
                f'border:1px solid {chip_fg};border-radius:10px;padding:3px 10px;'
                f'font-size:11px;margin:0 6px 6px 0;">{html_lib.escape(state)} '
                f'<b>{count}</b></span>'
            )

    # ON HOLD strip (bottom of slide): parked items, dimmed, two columns
    MAX_ONHOLD_SHOWN = 6
    onhold_cells = ""
    onhold_more = ""
    if onhold_count > 0 and number_col is not None:
        onhold_df = df[onhold_mask]
        shown = onhold_df.head(MAX_ONHOLD_SHOWN)
        items = []
        for _, row in shown.iterrows():
            desc = str(row[desc_col]) if desc_col is not None else ""
            if len(desc) > 44:
                desc = desc[:41] + "..."
            assignee = str(row[assigned_col]).strip() if assigned_col is not None else ""
            if not assignee:
                assignee = "UNASSIGNED"
            prio = str(row[priority_col]) if priority_col is not None else ""
            items.append(
                f'<td style="width:50%;padding:3px 10px 3px 0;font-size:11px;color:#777777;'
                f'border-bottom:1px dotted #D8D2C8;white-space:nowrap;overflow:hidden;">'
                f'<b style="color:#555555;">{html_lib.escape(str(row[number_col]))}</b>'
                f' &nbsp;<span style="font-size:10px;border:1px solid #AAAAAA;border-radius:8px;'
                f'padding:1px 6px;">{html_lib.escape(prio)}</span>'
                f' &nbsp;{html_lib.escape(desc)}'
                f' &nbsp;&mdash;&nbsp;<i>{html_lib.escape(assignee)}</i></td>'
            )
        rows_html = ""
        for i in range(0, len(items), 2):
            pair = items[i] + (items[i + 1] if i + 1 < len(items) else '<td style="width:50%;"></td>')
            rows_html = rows_html + "<tr>" + pair + "</tr>"
        onhold_cells = rows_html
        remaining = onhold_count - len(shown)
        if remaining > 0:
            onhold_more = (
                f'<span style="font-size:11px;color:#999999;">+ {remaining} more on hold '
                f'(see detailed report)</span>'
            )

    html_doc = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8" /><title>Executive Summary - ServiceNow Incidents</title></head>
<body style="margin:0;padding:0;background:#888888;font-family:Calibri,Arial,sans-serif;">
<div style="width:1280px;height:720px;background:{ANDE_CREAM};margin:0 auto;
     position:relative;overflow:hidden;box-sizing:border-box;padding:26px 34px;">

  <table cellpadding="0" cellspacing="0" style="width:100%;">
  <tr>
    <td style="vertical-align:middle;width:230px;">{build_wordmark_html(28)}</td>
    <td style="vertical-align:middle;padding-left:20px;border-left:3px solid {ANDE_GOLD};">
      <div style="font-size:23px;font-weight:bold;color:{ANDE_BLUE};line-height:1.1;">
        ServiceNow Incidents &mdash; Executive Summary</div>
      <div style="font-size:12px;color:#555555;margin-top:2px;">
        SN - Corporate Solutions &nbsp;|&nbsp; Open incidents (excl. Resolved) &nbsp;|&nbsp; {generated}</div>
    </td>
    <td style="text-align:right;vertical-align:middle;font-size:11px;color:{ANDE_BLUE};">
      TCS / Project Genesis<br/>Confidential</td>
  </tr>
  </table>
  <div style="border-bottom:3px solid {ANDE_GOLD};margin:12px 0 16px 0;"></div>

  <table cellpadding="0" cellspacing="0" style="width:100%;"><tr>{kpis}</tr></table>

  <table cellpadding="0" cellspacing="0" style="width:100%;margin-top:18px;">
  <tr>
    <td style="width:58%;vertical-align:top;padding-right:22px;">
      <div style="font-size:14px;font-weight:bold;color:{ANDE_BLUE};
           border-bottom:2px solid {ANDE_GOLD};padding-bottom:4px;margin-bottom:6px;">
        NEEDS ATTENTION &mdash; TOP ACTIVE INCIDENTS</div>
      <table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">
        {attention_rows if attention_rows else '<tr><td style="font-size:12px;color:#555;">No active incidents.</td></tr>'}
      </table>
    </td>
    <td style="width:42%;vertical-align:top;">
      <div style="font-size:14px;font-weight:bold;color:{ANDE_BLUE};
           border-bottom:2px solid {ANDE_GOLD};padding-bottom:4px;margin-bottom:6px;">
        OPEN WORKLOAD BY ASSIGNEE</div>
      <table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">
        {assignee_bars if assignee_bars else '<tr><td style="font-size:12px;color:#555;">No data.</td></tr>'}
      </table>
      <div style="font-size:14px;font-weight:bold;color:{ANDE_BLUE};
           border-bottom:2px solid {ANDE_GOLD};padding-bottom:4px;
           margin:14px 0 8px 0;">BY STATE</div>
      <div>{state_chips if state_chips else '<span style="font-size:12px;color:#555;">No data.</span>'}</div>
    </td>
  </tr>
  </table>

  <div style="margin-top:14px;">
    <div style="font-size:13px;font-weight:bold;color:#777777;
         border-bottom:2px solid {ANDE_GOLD};padding-bottom:3px;margin-bottom:4px;">
      ON HOLD &mdash; PARKED ITEMS (RANKED LAST)
      &nbsp;<span style="font-weight:normal;font-size:11px;color:#999999;">{onhold_count} total</span>
      &nbsp;{onhold_more}</div>
    <table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;table-layout:fixed;">
      {onhold_cells if onhold_cells else '<tr><td style="font-size:11px;color:#999999;padding:3px 0;">None - no incidents currently on hold.</td></tr>'}
    </table>
  </div>

  <div style="position:absolute;bottom:16px;left:34px;right:34px;
       border-top:2px solid {ANDE_GOLD};padding-top:7px;font-size:10.5px;color:{ANDE_BLUE};">
    <b>Harish Gupta</b> | Enterprise Solutions Unit &mdash; Consulting Practice | gupta.h@tcs.com
    <span style="float:right;color:#777777;">On Hold items ranked last regardless of priority
    &nbsp;&bull;&nbsp; Source: ServiceNow incident list extract</span>
  </div>
</div>
</body>
</html>"""
    output_file.write_text(html_doc, encoding="utf-8")


def write_pptx_slide(df: pd.DataFrame, output_file: Path) -> bool:
    """Native single-slide PowerPoint version of the executive dashboard.
    Returns False (with a console note) if python-pptx is not installed."""
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt, Emu
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
    except ImportError:
        print("Note: python-pptx not installed - skipping .pptx slide. Install with: pip install python-pptx")
        return False

    BLUE = RGBColor(0x00, 0x2D, 0x5B)
    GOLD = RGBColor(0xFF, 0xB4, 0x32)
    CREAM = RGBColor(0xFA, 0xF8, 0xF4)
    GRAY = RGBColor(0x77, 0x77, 0x77)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    DARK = RGBColor(0x33, 0x33, 0x33)
    PRIO_RGB = {1: RGBColor(0xC0, 0x00, 0x00), 2: RGBColor(0xC5, 0x5A, 0x11),
                3: RGBColor(0xBF, 0x8F, 0x00), 4: RGBColor(0x54, 0x82, 0x35),
                5: RGBColor(0x2E, 0x75, 0xB6)}

    priority_col = find_column(df.columns, "Priority")
    state_col = find_column(df.columns, "State", "Status")
    assigned_col = find_column(df.columns, "Assigned To", "Assigned to")
    number_col = find_column(df.columns, "Number", "Task number")
    desc_col = find_column(df.columns, "Short Description", "Short description")

    total = len(df)
    counts = {1: 0, 2: 0, 3: 0, 4: 0}
    if priority_col is not None:
        for value in df[priority_col]:
            rank = priority_rank_of(value)
            if rank in counts:
                counts[rank] += 1
    if state_col is not None and not df.empty:
        onhold_mask = df[state_col].astype(str).str.strip().str.lower().eq("on hold")
    else:
        onhold_mask = pd.Series([False] * len(df), index=df.index, dtype=bool)
    onhold_count = int(onhold_mask.sum())

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = CREAM; bg.line.fill.background()
    bg.shadow.inherit = False

    def textbox(left, top, width, height, text, size, color, bold=False, align=PP_ALIGN.LEFT):
        box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = "Calibri"
        return box

    textbox(0.45, 0.30, 3.3, 0.6, PROGRAM_NAME, 24, BLUE, bold=True)
    textbox(3.9, 0.26, 8.0, 0.5, "ServiceNow Incidents - Executive Summary", 24, BLUE, bold=True)
    textbox(3.9, 0.74, 8.0, 0.35,
            f"SN - Corporate Solutions  |  Open incidents (excl. Resolved)  |  "
            f"{datetime.now():%B %d, %Y}", 11, GRAY)
    rule = slide.shapes.add_shape(1, Inches(0.45), Inches(1.18), Inches(12.45), Emu(38100))
    rule.fill.solid(); rule.fill.fore_color.rgb = GOLD; rule.line.fill.background()
    rule.shadow.inherit = False

    kpi_specs = [
        (str(total), "TOTAL OPEN", BLUE),
        (str(counts[1]), "CRITICAL", PRIO_RGB[1]),
        (str(counts[2]), "HIGH", PRIO_RGB[2]),
        (str(counts[3]), "MEDIUM", PRIO_RGB[3]),
        (str(counts[4]), "LOW", PRIO_RGB[4]),
        (str(onhold_count), "ON HOLD", GRAY),
    ]
    box_w, gap, left0 = 1.95, 0.15, 0.45
    for i, (value, label, color) in enumerate(kpi_specs):
        left = left0 + i * (box_w + gap)
        shape = slide.shapes.add_shape(5, Inches(left), Inches(1.45), Inches(box_w), Inches(1.15))
        shape.fill.solid(); shape.fill.fore_color.rgb = WHITE
        shape.line.color.rgb = color
        shape.line.width = Pt(1.5)
        shape.shadow.inherit = False
        tf = shape.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]; p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run(); r1.text = value
        r1.font.size = Pt(30); r1.font.bold = True; r1.font.color.rgb = color
        p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run(); r2.text = label
        r2.font.size = Pt(10); r2.font.color.rgb = BLUE

    textbox(0.45, 2.85, 7.4, 0.35, "NEEDS ATTENTION - TOP ACTIVE INCIDENTS", 13, BLUE, bold=True)
    active = df[~onhold_mask].head(6) if not df.empty else df
    n_rows = max(1, len(active)) + 1
    table_shape = slide.shapes.add_table(
        n_rows, 4, Inches(0.45), Inches(3.25), Inches(7.4), Inches(0.32 * n_rows)
    )
    table = table_shape.table
    table.columns[0].width = Inches(1.35)
    table.columns[1].width = Inches(1.15)
    table.columns[2].width = Inches(3.35)
    table.columns[3].width = Inches(1.55)
    for c, header in enumerate(["Number", "Priority", "Short Description", "Assigned To"]):
        cell = table.cell(0, c)
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE
        p = cell.text_frame.paragraphs[0]
        run = p.add_run(); run.text = header
        run.font.size = Pt(10); run.font.bold = True; run.font.color.rgb = WHITE
    for r, (_, row) in enumerate(active.iterrows(), start=1):
        rank = priority_rank_of(row[priority_col]) if priority_col is not None else None
        desc = str(row[desc_col]) if desc_col is not None else ""
        if len(desc) > 52:
            desc = desc[:49] + "..."
        assignee = str(row[assigned_col]).strip() if assigned_col is not None else ""
        cells = [
            str(row[number_col]) if number_col is not None else "",
            str(row[priority_col]) if priority_col is not None else "",
            desc,
            assignee if assignee else "UNASSIGNED",
        ]
        for c, value in enumerate(cells):
            cell = table.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if r % 2 else CREAM
            p = cell.text_frame.paragraphs[0]
            run = p.add_run(); run.text = value
            run.font.size = Pt(9.5)
            run.font.color.rgb = PRIO_RGB.get(rank, BLUE) if c == 1 else DARK
            run.font.bold = c in (0, 1)

    textbox(8.15, 2.85, 4.7, 0.35, "OPEN WORKLOAD BY ASSIGNEE", 13, BLUE, bold=True)
    if assigned_col is not None and not df.empty:
        workload = (
            df[assigned_col].astype(str).str.strip().replace("", "UNASSIGNED")
            .value_counts().head(7)
        )
        max_count = int(workload.max()) if len(workload) else 1
        top = 3.30
        for name, count in workload.items():
            textbox(8.15, top, 1.9, 0.28, str(name)[:20], 9.5, DARK)
            bar_w = max(0.25, 2.2 * count / max_count)
            bar = slide.shapes.add_shape(1, Inches(10.1), Inches(top + 0.04),
                                         Inches(bar_w), Inches(0.16))
            bar.fill.solid(); bar.fill.fore_color.rgb = BLUE; bar.line.fill.background()
            bar.shadow.inherit = False
            textbox(10.15 + bar_w, top, 0.5, 0.28, str(count), 9.5, BLUE, bold=True)
            top += 0.31

    # ON HOLD parked items (bottom strip, dimmed)
    MAX_ONHOLD_SHOWN = 3
    onhold_df = df[onhold_mask] if not df.empty else df
    header_text = f"ON HOLD — PARKED ITEMS (RANKED LAST)   |   {onhold_count} total"
    remaining = onhold_count - min(onhold_count, MAX_ONHOLD_SHOWN)
    if remaining > 0:
        header_text += f"   |   + {remaining} more in detailed report"
    textbox(0.45, 5.62, 12.4, 0.3, header_text, 11.5, GRAY, bold=True)
    hold_rule = slide.shapes.add_shape(1, Inches(0.45), Inches(5.92), Inches(12.45), Emu(19050))
    hold_rule.fill.solid(); hold_rule.fill.fore_color.rgb = GOLD; hold_rule.line.fill.background()
    hold_rule.shadow.inherit = False
    if onhold_count > 0:
        top = 6.02
        for _, row in onhold_df.head(MAX_ONHOLD_SHOWN).iterrows():
            desc = str(row[desc_col]) if desc_col is not None else ""
            if len(desc) > 70:
                desc = desc[:67] + "..."
            assignee = str(row[assigned_col]).strip() if assigned_col is not None else ""
            line = (
                f"{row[number_col] if number_col is not None else ''}   "
                f"[{row[priority_col] if priority_col is not None else ''}]   "
                f"{desc}   —  {assignee if assignee else 'UNASSIGNED'}"
            )
            textbox(0.45, top, 12.4, 0.27, line, 9, GRAY)
            top += 0.27
    else:
        textbox(0.45, 6.02, 12.4, 0.27, "None - no incidents currently on hold.", 9, GRAY)

    foot = slide.shapes.add_shape(1, Inches(0.45), Inches(6.92), Inches(12.45), Emu(25400))
    foot.fill.solid(); foot.fill.fore_color.rgb = GOLD; foot.line.fill.background()
    foot.shadow.inherit = False
    textbox(0.45, 7.02, 8.0, 0.35,
            "Harish Gupta | Enterprise Solutions Unit - Consulting Practice | gupta.h@tcs.com",
            9, BLUE)
    textbox(8.0, 7.02, 4.9, 0.35,
            "On Hold ranked last regardless of priority  |  TCS / Project Genesis Confidential",
            8.5, GRAY, align=PP_ALIGN.RIGHT)

    prs.save(str(output_file))
    return True


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
                print(f"Branded HTML report saved to: {html_file.resolve()}")
            except Exception as exc:
                print(f"HTML report generation failed ({exc}). Excel output is unaffected.")

            dashboard_file = output_base / (output_file.stem + "_dashboard.html")
            try:
                write_dashboard_html(filtered_df, dashboard_file)
                print(f"Executive dashboard (16:9 slide) saved to: {dashboard_file.resolve()}")
            except Exception as exc:
                print(f"Dashboard generation failed ({exc}). Other outputs are unaffected.")

            pptx_file = output_base / (output_file.stem + "_summary.pptx")
            try:
                if write_pptx_slide(filtered_df, pptx_file):
                    print(f"PowerPoint summary slide saved to: {pptx_file.resolve()}")
            except Exception as exc:
                print(f"PPTX slide generation failed ({exc}). Other outputs are unaffected.")

            print(f"Filtered rows exported: {len(filtered_df)}")

            browser.close()
            return 0

    except Exception as exc:
        print(f"Extraction failed: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
