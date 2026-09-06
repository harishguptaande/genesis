"""Unit tests for ServiceNow report helpers (no browser / network)."""

from __future__ import annotations

import pandas as pd
import pytest

from scripts.servicenow_incident_report import (
    PROGRAM_NAME,
    drop_placeholder_columns,
    filter_and_sort_records,
    find_column,
    normalize_text,
    priority_rank_of,
    write_dashboard_html,
    write_html_report,
    write_pptx_slide,
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


def test_drop_placeholder_columns_removes_column_n_even_with_values() -> None:
    df = pd.DataFrame(
        {
            "Number": ["INC1"],
            "Column_1": ["x"],
            "Priority": ["2 - High"],
            "Empty": [""],
        }
    )
    cleaned = drop_placeholder_columns(df)
    assert list(cleaned.columns) == ["Number", "Priority"]


def test_filter_and_sort_records_on_hold_last_and_excludes_resolved() -> None:
    records = [
        {
            "Number": "INC_HOLD",
            "Assignment group": "SN - Corporate Solutions",
            "State": "On Hold",
            "Priority": "1 - Critical",
        },
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
    assert list(df["Number"]) == ["INC1", "INC3", "INC_HOLD"]


def test_write_html_report_uses_project_genesis_wordmark(tmp_path) -> None:
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
    assert PROGRAM_NAME in html
    assert "ServiceNow Open Incident Report" in html
    assert "INC1" in html
    assert "The Andersons" not in html
    assert "<script" not in html.lower()


def test_write_dashboard_html_includes_kpis_and_on_hold(tmp_path) -> None:
    df = pd.DataFrame(
        [
            {
                "Number": "INC1",
                "Priority": "1 - Critical",
                "Short description": "VPN down",
                "State": "New",
                "Assigned to": "Alice",
            },
            {
                "Number": "INC9",
                "Priority": "2 - High",
                "Short description": "Waiting on vendor",
                "State": "On Hold",
                "Assigned to": "Bob",
            },
        ]
    )
    out = tmp_path / "dashboard.html"
    write_dashboard_html(df, out)
    html = out.read_text(encoding="utf-8")
    assert "Executive Summary" in html
    assert "1280px" in html
    assert "INC1" in html
    assert "ON HOLD" in html
    assert "INC9" in html
    assert PROGRAM_NAME in html


def test_write_pptx_slide_creates_file_when_python_pptx_available(tmp_path) -> None:
    pytest.importorskip("pptx")

    df = pd.DataFrame(
        [
            {
                "Number": "INC1",
                "Priority": "1 - Critical",
                "Short description": "VPN down",
                "State": "New",
                "Assigned to": "Alice",
            }
        ]
    )
    out = tmp_path / "summary.pptx"
    assert write_pptx_slide(df, out) is True
    assert out.is_file()
    assert out.stat().st_size > 0
