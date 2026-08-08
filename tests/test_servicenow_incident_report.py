"""Unit tests for ServiceNow report helpers (no browser / network)."""

from __future__ import annotations

import pandas as pd

from scripts.servicenow_incident_report import (
    drop_placeholder_columns,
    filter_and_sort_records,
    find_column,
    normalize_text,
    priority_rank_of,
    write_html_report,
)


def test_normalize_text_collapses_whitespace_and_nbsp() -> None:
    assert normalize_text("  foo\xa0\n bar  ") == "foo bar"


def test_find_column_is_case_insensitive() -> None:
    assert find_column(["Priority", "State"], "priority") == "Priority"
    assert find_column(["Assignment group"], "Assigned to group") is None
    assert find_column(["State"], "State", "Status") == "State"


def test_priority_rank_of_common_labels() -> None:
    assert priority_rank_of("1 - Critical") == 1
    assert priority_rank_of("High") == 2
    assert priority_rank_of("3 - Medium") == 3
    assert priority_rank_of("unknown") is None


def test_drop_placeholder_columns_removes_all_empty() -> None:
    df = pd.DataFrame(
        {
            "Number": ["INC1"],
            "Column_1": [""],
            "Priority": ["2 - High"],
        }
    )
    cleaned = drop_placeholder_columns(df)
    assert list(cleaned.columns) == ["Number", "Priority"]


def test_filter_and_sort_records_filters_and_orders() -> None:
    records = [
        {
            "Number": "INC3",
            "Assignment group": "SN - Corporate Solutions",
            "State": "In Progress",
            "Priority": "3 - Moderate",
        },
        {
            "Number": "INC1",
            "Assignment group": "SN - Corporate Solutions",
            "State": "New",
            "Priority": "1 - Critical",
        },
        {
            "Number": "INC2",
            "Assignment group": "SN - Corporate Solutions",
            "State": "Resolved",
            "Priority": "2 - High",
        },
        {
            "Number": "INC4",
            "Assignment group": "Other Group",
            "State": "New",
            "Priority": "1 - Critical",
        },
    ]
    df = filter_and_sort_records(records)
    assert list(df["Number"]) == ["INC1", "INC3"]


def test_write_html_report_is_self_contained(tmp_path) -> None:
    df = pd.DataFrame(
        [
            {
                "Number": "INC1",
                "Priority": "1 - Critical",
                "Short description": "VPN down",
                "State": "New",
            }
        ]
    )
    out = tmp_path / "report.html"
    write_html_report(df, out)
    html = out.read_text(encoding="utf-8")
    assert "ServiceNow Open Incident Report" in html
    assert "INC1" in html
    assert "The Andersons" in html
    assert "Critical" in html
    assert "<script" not in html.lower()
