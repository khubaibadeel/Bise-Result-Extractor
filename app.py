#!/usr/bin/env python3
"""BISE Faisalabad public result extractor.

Reads a user-supplied roll-number range, requests the current public
ResultDetails.aspx pages, extracts the student's basic details, and writes
Excel + HTML reports.

The scraper intentionally does not attempt to bypass CAPTCHAs, authentication,
or access controls. If the board blocks automated requests, the run records the
failure and stops/continues safely rather than trying to evade the block.
"""

from __future__ import annotations

import asyncio
import html
import random
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import aiohttp
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


BASE_URL = "https://result.bisefsd.edu.pk/ResultDetails.aspx?RollNo={}"
HOME_URL = "https://result.bisefsd.edu.pk/"
MIN_ROLL_LENGTH = 6
MAX_ROLL_LENGTH = 6
MAX_RANGE_SIZE = 1000000
MAX_CONCURRENT_REQUESTS = 500
REQUEST_TIMEOUT_SECONDS = 20
MAX_RETRIES = 3


@dataclass
class StudentResult:
    requested_roll: str
    roll_number: str = ""
    full_name: str = ""
    father_name: str = ""
    total_marks: Optional[int] = None
    notification: str = ""
    academic_status: str = "unknown"
    fetch_status: str = ""
    error: str = ""


def clean_text(value: str) -> str:
    """Collapse whitespace and strip surrounding space."""
    return re.sub(r"\s+", " ", value or "").strip()


def normalize_label(value: str) -> str:
    """Normalize labels so small punctuation/case differences do not matter."""
    value = clean_text(value).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return clean_text(value)


def parse_integer(value: str) -> Optional[int]:
    """Parse a plain non-negative integer such as '1,042'."""
    cleaned = clean_text(value).replace(",", "")
    if re.fullmatch(r"\d+", cleaned):
        return int(cleaned)
    return None


def classify_academic_status(notification: str, total_marks: Optional[int]) -> str:
    """Classify the board notification without treating all text as supply."""
    text = normalize_label(notification)

    if not text:
        return "regular" if total_marks is not None else "unknown"

    supply_words = (
        "supply",
        "supplementary",
        "re appear",
        "reappear",
        "compartment",
    )
    fail_words = ("fail", "failed")
    absent_words = ("absent", "abs")
    withheld_words = (
        "withheld",
        "result later",
        "rl",
        "cancelled",
        "cancel",
        "unfair means",
        "ufm",
    )
    pass_words = ("pass", "passed", "qualified", "successful")

    if any(word in text for word in supply_words):
        return "supply"
    if any(word in text for word in fail_words):
        return "failed"
    if any(word in text for word in absent_words):
        return "absent"
    if any(word in text for word in withheld_words):
        return "withheld"
    if any(word in text for word in pass_words):
        return "regular"

    # On some board pages the fourth info value is simply the numeric total.
    if parse_integer(notification) is not None:
        return "regular"

    return "other"


def extract_labeled_info(soup: BeautifulSoup) -> dict[str, str]:
    """Extract label/value pairs from table, definition-list and inline layouts.

    The 2026 ResultDetails.aspx page renders important fields as visible text
    such as ``Roll No. : 473428`` and ``Notification: 1134`` rather than the
    older ``.info > b`` layout, so this function supports both structures.
    """
    values: dict[str, str] = {}

    # Table-like layouts: <tr><td>Roll No</td><td>...</td></tr>
    for row in soup.find_all("tr"):
        cells = row.find_all(["th", "td"], recursive=False)
        if len(cells) < 2:
            continue
        label = normalize_label(cells[0].get_text(" ", strip=True))
        value = clean_text(cells[1].get_text(" ", strip=True))
        if label and value:
            values.setdefault(label, value)

    # Definition-list layouts.
    for dt in soup.find_all("dt"):
        dd = dt.find_next_sibling("dd")
        if dd is None:
            continue
        label = normalize_label(dt.get_text(" ", strip=True))
        value = clean_text(dd.get_text(" ", strip=True))
        if label and value:
            values.setdefault(label, value)

    # Current 2026 page: labels and values are often emitted as ordinary text.
    lines = [clean_text(text) for text in soup.stripped_strings]
    aliases = {
        "roll no": ("roll no", "roll number"),
        "candidate name": ("name of candidate", "candidate name", "student name"),
        "father name": ("father s name", "fathers name", "father name"),
        "notification": ("notification",),
    }

    for canonical, label_aliases in aliases.items():
        if canonical in values:
            continue
        for index, line in enumerate(lines):
            normalized = normalize_label(line)
            matched_alias = next(
                (alias for alias in label_aliases if normalized.startswith(alias)),
                None,
            )
            if matched_alias is None:
                continue

            # Prefer text after ':' from the same visible line.
            if ":" in line:
                remainder = clean_text(line.split(":", 1)[1])
                if remainder:
                    values[canonical] = remainder
                    break

            # Some markup places the label and value in adjacent text nodes.
            normalized_alias = normalize_label(matched_alias)
            if normalized == normalized_alias and index + 1 < len(lines):
                candidate = clean_text(lines[index + 1])
                if candidate:
                    values[canonical] = candidate
                    break

            # Last fallback: remove the matched label from the beginning.
            pattern = re.compile(r"^" + re.escape(matched_alias).replace(r"\ ", r"\s+") + r"\s*[.:\-]*\s*", re.I)
            remainder = clean_text(pattern.sub("", normalized, count=1))
            if remainder and remainder != normalized:
                values[canonical] = remainder
                break


    return values

def first_matching_value(info: dict[str, str], aliases: tuple[str, ...]) -> str:
    """Return the first value whose normalized label matches an alias."""
    normalized_aliases = tuple(normalize_label(alias) for alias in aliases)

    for alias in normalized_aliases:
        if alias in info:
            return info[alias]

    # Permit labels with harmless suffixes/prefixes such as ':' or '(candidate)'.
    for label, value in info.items():
        if any(alias and (alias in label or label in alias) for alias in normalized_aliases):
            return value

    return ""


def extract_info_bold_values(soup: BeautifulSoup) -> list[str]:
    """Fallback for the board's compact `.info` block."""
    info = soup.find("div", class_=lambda value: value and "info" in value.split())
    if info is None:
        return []

    return [
        clean_text(element.get_text(" ", strip=True))
        for element in info.find_all("b")
        if clean_text(element.get_text(" ", strip=True))
    ]


def find_subject_marks_column(table) -> Optional[int]:
    """Find the subject-total column used by the current BISE result table.

    On the 2026 Faisalabad page the right-most column is labelled ``Remarks``
    but, for ordinary results, it contains each subject's obtained total.
    We still detect the column by header instead of blindly assuming the last
    cell so layout changes fail safely.
    """
    for row in table.find_all("tr")[:4]:
        cells = row.find_all(["th", "td"])
        for index, cell in enumerate(cells):
            label = normalize_label(cell.get_text(" ", strip=True))
            if label == "remarks":
                return index
            if "obtained" in label and "mark" in label:
                return index
            if label in {"obtained", "marks obtained", "total obtained"}:
                return index
    return None


def calculate_total_from_subjects(soup: BeautifulSoup) -> Optional[int]:
    """Sum numeric per-subject totals from the detected result-table column."""
    for table in soup.find_all("table"):
        marks_index = find_subject_marks_column(table)
        if marks_index is None:
            continue

        total = 0
        valid_rows = 0
        for row in table.find_all("tr"):
            cells = row.find_all("td")
            if len(cells) <= marks_index:
                continue

            value = parse_integer(cells[marks_index].get_text(" ", strip=True))
            if value is None or value > 300:
                continue

            total += value
            valid_rows += 1

        # The SSC page has several subject rows; requiring at least 3 avoids
        # accidentally summing unrelated small tables.
        if valid_rows >= 3:
            return total

    return None

def page_says_not_found(soup: BeautifulSoup, html_text: str) -> bool:
    """Recognize common 'no result' responses without relying on one title."""
    title = clean_text(soup.title.get_text(" ", strip=True)) if soup.title else ""
    visible_text = normalize_label(soup.get_text(" ", strip=True)[:4000])
    title_text = normalize_label(title)

    phrases = (
        "result not found",
        "record not found",
        "no result found",
        "no record found",
        "invalid roll number",
    )

    return any(phrase in title_text or phrase in visible_text for phrase in phrases)


def parse_result(html_text: str, requested_roll: str) -> StudentResult:
    """Parse one current BISE Faisalabad ResultDetails.aspx page."""
    soup = BeautifulSoup(html_text, "html.parser")

    if page_says_not_found(soup, html_text):
        return StudentResult(requested_roll=requested_roll, fetch_status="not_found")

    labeled = extract_labeled_info(soup)

    roll_number = first_matching_value(
        labeled, ("roll no", "roll number", "roll #", "roll")
    )
    full_name = first_matching_value(
        labeled, ("candidate name", "name of candidate", "student name", "name")
    )
    father_name = first_matching_value(
        labeled, ("father name", "father's name", "fathers name")
    )
    notification = first_matching_value(labeled, ("notification",))

    # Older cached layouts used `.info > b`; retain it only as a fallback.
    if not (roll_number and full_name and father_name):
        bold_values = extract_info_bold_values(soup)
        if len(bold_values) >= 3:
            roll_number = roll_number or bold_values[0]
            full_name = full_name or bold_values[1]
            father_name = father_name or bold_values[2]
            if len(bold_values) >= 4:
                notification = notification or bold_values[3]

    # Roll number is the strongest sanity check against an unexpected redirect
    # or generic landing page returned with HTTP 200.
    parsed_roll_match = re.search(r"\d+", roll_number or "")
    parsed_roll = parsed_roll_match.group(0) if parsed_roll_match else ""

    if not parsed_roll or not full_name or not father_name:
        return StudentResult(
            requested_roll=requested_roll,
            roll_number=parsed_roll,
            full_name=full_name,
            father_name=father_name,
            notification=notification,
            fetch_status="parse_error",
            error=(
                "The server returned HTTP 200, but the current result fields "
                "could not be identified. The page may be a landing/error page "
                "or the BISE layout may have changed."
            ),
        )

    if parsed_roll.lstrip("0") != requested_roll.lstrip("0"):
        return StudentResult(
            requested_roll=requested_roll,
            roll_number=parsed_roll,
            full_name=full_name,
            father_name=father_name,
            notification=notification,
            fetch_status="parse_error",
            error=f"Requested roll {requested_roll}, but page contains roll {parsed_roll}.",
        )

    total_marks = parse_integer(notification)
    if total_marks is None:
        total_marks = calculate_total_from_subjects(soup)

    academic_status = classify_academic_status(notification, total_marks)

    return StudentResult(
        requested_roll=requested_roll,
        roll_number=parsed_roll,
        full_name=full_name,
        father_name=father_name,
        total_marks=total_marks,
        notification=notification,
        academic_status=academic_status,
        fetch_status="found",
    )


async def fetch_one(
    session: aiohttp.ClientSession,
    semaphore: asyncio.Semaphore,
    roll_number: str,
) -> StudentResult:
    """Fetch one result with bounded concurrency and retryable errors."""
    url = BASE_URL.format(roll_number)

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            async with semaphore:
                async with session.get(url, allow_redirects=True) as response:
                    status = response.status

                    if status == 404:
                        return StudentResult(
                            requested_roll=roll_number,
                            fetch_status="not_found",
                        )

                    if status == 403:
                        return StudentResult(
                            requested_roll=roll_number,
                            fetch_status="blocked",
                            error="HTTP 403: access denied by the result server.",
                        )

                    if status == 429:
                        if attempt == MAX_RETRIES:
                            return StudentResult(
                                requested_roll=roll_number,
                                fetch_status="rate_limited",
                                error="HTTP 429: too many requests.",
                            )
                        retry_after = response.headers.get("Retry-After")
                        try:
                            delay = float(retry_after) if retry_after else 2 ** attempt
                        except ValueError:
                            delay = 2 ** attempt
                        delay += random.uniform(0.2, 0.8)
                    elif status >= 500:
                        if attempt == MAX_RETRIES:
                            return StudentResult(
                                requested_roll=roll_number,
                                fetch_status="server_error",
                                error=f"HTTP {status}: server error.",
                            )
                        delay = 2 ** (attempt - 1) + random.uniform(0.2, 0.8)
                    elif status != 200:
                        return StudentResult(
                            requested_roll=roll_number,
                            fetch_status="http_error",
                            error=f"HTTP {status}",
                        )
                    else:
                        page_html = await response.text(errors="replace")
                        try:
                            return parse_result(page_html, roll_number)
                        except Exception as error:  # parser must not kill the whole batch
                            return StudentResult(
                                requested_roll=roll_number,
                                fetch_status="parse_error",
                                error=f"Parser error: {error}",
                            )

            await asyncio.sleep(delay)

        except (aiohttp.ClientError, asyncio.TimeoutError) as error:
            if attempt == MAX_RETRIES:
                return StudentResult(
                    requested_roll=roll_number,
                    fetch_status="network_error",
                    error=f"{type(error).__name__}: {error}",
                )

            delay = 2 ** (attempt - 1) + random.uniform(0.2, 0.8)
            await asyncio.sleep(delay)

    return StudentResult(
        requested_roll=roll_number,
        fetch_status="error",
        error="Unexpected request failure.",
    )


def validate_range(start: str, end: str) -> tuple[str, str]:
    start = start.strip()
    end = end.strip()

    if not start.isdigit() or not end.isdigit():
        raise ValueError("Roll numbers must contain digits only.")

    if len(start) != len(end):
        raise ValueError("Starting and ending roll numbers must have the same length.")

    if not (MIN_ROLL_LENGTH <= len(start) <= MAX_ROLL_LENGTH):
        raise ValueError(
            f"Roll numbers must be between {MIN_ROLL_LENGTH} and "
            f"{MAX_ROLL_LENGTH} digits long."
        )

    start_number = int(start)
    end_number = int(end)

    if start_number > end_number:
        raise ValueError("Starting roll number cannot exceed ending roll number.")

    range_size = end_number - start_number + 1
    if range_size > MAX_RANGE_SIZE:
        raise ValueError(
            f"Range contains {range_size:,} roll numbers; maximum per run is "
            f"{MAX_RANGE_SIZE:,}. Split it into smaller batches."
        )

    return start, end


def is_supply_result(result: StudentResult) -> bool:
    return result.fetch_status == "found" and result.academic_status == "supply"


def save_to_excel(results: list[StudentResult], output_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Results"

    headers = [
        "Roll Number",
        "Full Name",
        "Father Name",
        "Total Marks",
        "Academic Status",
        "Notification",
    ]
    worksheet.append(headers)

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    found_results = sorted(
        (result for result in results if result.fetch_status == "found"),
        key=lambda item: int(item.requested_roll),
    )

    for result in found_results:
        worksheet.append([
            result.roll_number,
            result.full_name,
            result.father_name,
            result.total_marks,
            result.academic_status.title(),
            result.notification,
        ])

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    widths = {1: 16, 2: 35, 3: 35, 4: 15, 5: 18, 6: 50}
    for column_number, width in widths.items():
        worksheet.column_dimensions[get_column_letter(column_number)].width = width

    for cell in worksheet["A"][1:]:
        cell.number_format = "@"

    log_sheet = workbook.create_sheet("Log")
    log_sheet.append(["Requested Roll", "Fetch Status", "Error"])
    for cell in log_sheet[1]:
        cell.font = Font(bold=True)

    for result in sorted(results, key=lambda item: int(item.requested_roll)):
        if result.fetch_status == "found":
            continue
        log_sheet.append([result.requested_roll, result.fetch_status, result.error])

    log_sheet.freeze_panes = "A2"
    log_sheet.auto_filter.ref = log_sheet.dimensions
    log_sheet.column_dimensions["A"].width = 18
    log_sheet.column_dimensions["B"].width = 20
    log_sheet.column_dimensions["C"].width = 85

    workbook.save(output_path)


def status_label(status: str) -> str:
    labels = {
        "regular": "Pass / Regular",
        "supply": "Supply",
        "failed": "Failed",
        "absent": "Absent",
        "withheld": "Withheld / RL",
        "other": "Other",
        "unknown": "Unknown",
    }
    return labels.get(status, status.title())


def save_to_html(
    results: list[StudentResult],
    output_path: Path,
    start_roll: str,
    end_roll: str,
) -> None:
    found_results = sorted(
        (result for result in results if result.fetch_status == "found"),
        key=lambda item: int(item.requested_roll),
    )

    supply_count = sum(result.academic_status == "supply" for result in found_results)
    non_regular_count = sum(
        result.academic_status not in {"regular", "unknown"}
        for result in found_results
    )
    marks_values = [
        result.total_marks for result in found_results if result.total_marks is not None
    ]
    highest_marks = max(marks_values) if marks_values else "—"

    rows: list[str] = []
    for result in found_results:
        row_status = result.academic_status
        marks_text = str(result.total_marks) if result.total_marks is not None else "—"

        rows.append(
            f"""
            <tr
                data-roll="{html.escape(result.roll_number, quote=True)}"
                data-name="{html.escape(result.full_name.lower(), quote=True)}"
                data-father="{html.escape(result.father_name.lower(), quote=True)}"
                data-status="{html.escape(row_status, quote=True)}"
            >
                <td class="roll">{html.escape(result.roll_number)}</td>
                <td>{html.escape(result.full_name)}</td>
                <td>{html.escape(result.father_name)}</td>
                <td class="marks">{html.escape(marks_text)}</td>
                <td><span class="badge {html.escape(row_status, quote=True)}">{html.escape(status_label(row_status))}</span></td>
                <td class="notification">{html.escape(result.notification or "—")}</td>
            </tr>
            """
        )

    page = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>BISE Faisalabad Results {html.escape(start_roll)}–{html.escape(end_roll)}</title>
    <style>
        :root {{
            --bg: #f4f7fb;
            --panel: #ffffff;
            --text: #172033;
            --muted: #667085;
            --border: #dfe5ee;
            --accent: #006f83;
            --accent-dark: #004f5e;
            --good-bg: #e9f8ef;
            --good-text: #1f7a45;
            --warn-bg: #fff3d9;
            --warn-text: #946200;
            --bad-bg: #fdecec;
            --bad-text: #b42318;
            --neutral-bg: #eef2f6;
            --neutral-text: #475467;
            --shadow: 0 12px 32px rgba(20, 42, 70, 0.09);
        }}
        * {{ box-sizing: border-box; }}
        body {{
            margin: 0;
            background: var(--bg);
            color: var(--text);
            font-family: Inter, "Segoe UI", Arial, sans-serif;
        }}
        .hero {{
            padding: 38px 20px 90px;
            background: radial-gradient(circle at top right, rgba(255,255,255,.16), transparent 30%),
                        linear-gradient(135deg, var(--accent-dark), var(--accent));
            color: white;
        }}
        .container {{ width: min(1200px, calc(100% - 28px)); margin: 0 auto; }}
        .hero h1 {{ margin: 0; font-size: clamp(1.75rem, 4vw, 2.75rem); }}
        .hero p {{ margin: 10px 0 0; color: rgba(255,255,255,.82); }}
        .dashboard {{ margin-top: -58px; padding-bottom: 40px; }}
        .stats {{
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 14px;
            margin-bottom: 18px;
        }}
        .card {{
            background: var(--panel);
            border: 1px solid var(--border);
            border-radius: 16px;
            box-shadow: var(--shadow);
        }}
        .stat {{ padding: 20px; }}
        .stat .label {{ color: var(--muted); font-size: .88rem; margin-bottom: 7px; }}
        .stat .value {{ font-size: 1.8rem; font-weight: 750; }}
        .controls {{
            display: grid;
            grid-template-columns: minmax(220px, 1fr) minmax(190px, auto) auto;
            gap: 12px;
            padding: 16px;
            margin-bottom: 18px;
        }}
        input, select, button {{
            min-height: 44px;
            border-radius: 10px;
            border: 1px solid var(--border);
            background: white;
            color: var(--text);
            font: inherit;
        }}
        input {{ width: 100%; padding: 0 14px; }}
        select, button {{ padding: 0 14px; }}
        button {{ cursor: pointer; background: var(--accent); border-color: var(--accent); color: white; font-weight: 650; }}
        button:hover {{ background: var(--accent-dark); }}
        .table-card {{ overflow: hidden; }}
        .table-wrap {{ overflow-x: auto; }}
        table {{ width: 100%; border-collapse: collapse; min-width: 920px; }}
        thead {{ background: #f8fafc; }}
        th, td {{ padding: 14px 16px; border-bottom: 1px solid var(--border); text-align: left; vertical-align: middle; }}
        th {{ color: #475467; font-size: .82rem; text-transform: uppercase; letter-spacing: .035em; cursor: pointer; user-select: none; }}
        tbody tr:hover {{ background: #f8fbfc; }}
        .roll, .marks {{ font-weight: 750; }}
        .roll {{ color: var(--accent-dark); }}
        .notification {{ color: var(--muted); max-width: 340px; }}
        .badge {{ display: inline-flex; align-items: center; justify-content: center; padding: 5px 10px; border-radius: 999px; font-weight: 700; font-size: .78rem; }}
        .badge.regular {{ background: var(--good-bg); color: var(--good-text); }}
        .badge.supply {{ background: var(--warn-bg); color: var(--warn-text); }}
        .badge.failed, .badge.absent {{ background: var(--bad-bg); color: var(--bad-text); }}
        .badge.withheld, .badge.other, .badge.unknown {{ background: var(--neutral-bg); color: var(--neutral-text); }}
        .footer-row {{ display: flex; justify-content: space-between; gap: 14px; align-items: center; padding: 14px 16px; color: var(--muted); font-size: .9rem; }}
        .empty {{ display: none; padding: 38px 20px; text-align: center; color: var(--muted); }}
        @media (max-width: 850px) {{ .stats {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} .controls {{ grid-template-columns: 1fr; }} }}
        @media (max-width: 500px) {{ .stats {{ grid-template-columns: 1fr; }} .hero {{ padding-top: 28px; }} }}
        @media print {{ .hero, .controls, .footer-row {{ display: none; }} .dashboard {{ margin-top: 0; }} body {{ background: white; }} .card {{ box-shadow: none; }} }}
    </style>
</head>
<body>
    <header class="hero">
        <div class="container">
            <h1>BISE Faisalabad Results</h1>
            <p>Roll-number range {html.escape(start_roll)} to {html.escape(end_roll)}</p>
        </div>
    </header>

    <main class="container dashboard">
        <section class="stats">
            <article class="card stat"><div class="label">Students found</div><div class="value">{len(found_results)}</div></article>
            <article class="card stat"><div class="label">Supply cases</div><div class="value">{supply_count}</div></article>
            <article class="card stat"><div class="label">Other non-regular</div><div class="value">{non_regular_count}</div></article>
            <article class="card stat"><div class="label">Highest marks</div><div class="value">{highest_marks}</div></article>
        </section>

        <section class="card controls">
            <input id="searchInput" type="search" placeholder="Search roll number, student, or father name" autocomplete="off">
            <select id="statusFilter">
                <option value="all">All statuses</option>
                <option value="regular">Pass / Regular</option>
                <option value="supply">Supply</option>
                <option value="failed">Failed</option>
                <option value="absent">Absent</option>
                <option value="withheld">Withheld / RL</option>
                <option value="other">Other</option>
                <option value="unknown">Unknown</option>
            </select>
            <button id="resetButton" type="button">Reset</button>
        </section>

        <section class="card table-card">
            <div class="table-wrap">
                <table id="resultsTable">
                    <thead><tr>
                        <th data-column="0">Roll Number</th>
                        <th data-column="1">Full Name</th>
                        <th data-column="2">Father Name</th>
                        <th data-column="3">Total Marks</th>
                        <th data-column="4">Status</th>
                        <th data-column="5">Notification</th>
                    </tr></thead>
                    <tbody>{''.join(rows)}</tbody>
                </table>
            </div>
            <div id="emptyState" class="empty">No students match the current search or filter.</div>
            <div class="footer-row">
                <span id="visibleCount">Showing {len(found_results)} students</span>
                <span>Generated locally by the BISE FSD extractor</span>
            </div>
        </section>
    </main>

    <script>
        const table = document.getElementById("resultsTable");
        const tbody = table.querySelector("tbody");
        const rowElements = Array.from(tbody.querySelectorAll("tr"));
        const searchInput = document.getElementById("searchInput");
        const statusFilter = document.getElementById("statusFilter");
        const resetButton = document.getElementById("resetButton");
        const visibleCount = document.getElementById("visibleCount");
        const emptyState = document.getElementById("emptyState");
        let sortColumn = 0;
        let sortAscending = true;

        function applyFilters() {{
            const query = searchInput.value.trim().toLowerCase();
            const wantedStatus = statusFilter.value;
            let shown = 0;

            rowElements.forEach(row => {{
                const searchable = [row.dataset.roll, row.dataset.name, row.dataset.father].join(" ").toLowerCase();
                const queryMatches = !query || searchable.includes(query);
                const statusMatches = wantedStatus === "all" || row.dataset.status === wantedStatus;
                const visible = queryMatches && statusMatches;
                row.hidden = !visible;
                if (visible) shown += 1;
            }});

            visibleCount.textContent = `Showing ${{shown}} student${{shown === 1 ? "" : "s"}}`;
            emptyState.style.display = shown ? "none" : "block";
        }}

        function numericCellValue(text) {{
            const cleaned = text.replace(/[^0-9.-]/g, "");
            if (!cleaned || cleaned === "-") return Number.NEGATIVE_INFINITY;
            const number = Number(cleaned);
            return Number.isFinite(number) ? number : Number.NEGATIVE_INFINITY;
        }}

        function compareRows(a, b, column) {{
            const aCell = a.children[column].textContent.trim();
            const bCell = b.children[column].textContent.trim();

            if (column === 0 || column === 3) {{
                return numericCellValue(aCell) - numericCellValue(bCell);
            }}

            return aCell.localeCompare(bCell, undefined, {{ sensitivity: "base", numeric: true }});
        }}

        table.querySelectorAll("th[data-column]").forEach(header => {{
            header.addEventListener("click", () => {{
                const column = Number(header.dataset.column);
                if (sortColumn === column) {{
                    sortAscending = !sortAscending;
                }} else {{
                    sortColumn = column;
                    sortAscending = true;
                }}

                const sorted = [...rowElements].sort((a, b) => {{
                    const comparison = compareRows(a, b, column);
                    return sortAscending ? comparison : -comparison;
                }});
                sorted.forEach(row => tbody.appendChild(row));
            }});
        }});

        searchInput.addEventListener("input", applyFilters);
        statusFilter.addEventListener("change", applyFilters);
        resetButton.addEventListener("click", () => {{
            searchInput.value = "";
            statusFilter.value = "all";
            applyFilters();
            searchInput.focus();
        }});

        applyFilters();
    </script>
</body>
</html>
"""

    output_path.write_text(page, encoding="utf-8")


async def run_scraper(start: str, end: str) -> list[StudentResult]:
    width = len(start)
    rolls = [
        str(number).zfill(width)
        for number in range(int(start), int(end) + 1)
    ]

    semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)
    timeout = aiohttp.ClientTimeout(total=REQUEST_TIMEOUT_SECONDS)
    connector = aiohttp.TCPConnector(
        limit=MAX_CONCURRENT_REQUESTS,
        ttl_dns_cache=300,
    )

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/151.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.8",
        "Referer": "https://result.bisefsd.edu.pk/",
    }

    async with aiohttp.ClientSession(
        headers=headers,
        timeout=timeout,
        connector=connector,
    ) as session:
        # Establish normal session cookies first. Failure here is non-fatal;
        # direct ResultDetails.aspx requests may still work.
        try:
            async with session.get(HOME_URL, allow_redirects=True) as response:
                await response.read()
        except (aiohttp.ClientError, asyncio.TimeoutError):
            pass

        tasks = [
            asyncio.create_task(fetch_one(session, semaphore, roll_number))
            for roll_number in rolls
        ]

        results: list[StudentResult] = []
        counters = {
            "found": 0,
            "supply": 0,
            "not_found": 0,
            "errors": 0,
        }

        for completed, task in enumerate(asyncio.as_completed(tasks), start=1):
            try:
                result = await task
            except Exception as error:  # last-resort protection for one failed task
                counters["errors"] += 1
                print(f"[{completed}/{len(rolls)}] unexpected task error: {error}")
                continue

            results.append(result)

            if result.fetch_status == "found":
                counters["found"] += 1
                if result.academic_status == "supply":
                    counters["supply"] += 1

                marks = result.total_marks if result.total_marks is not None else "N/A"
                print(
                    f"[{completed}/{len(rolls)}] {result.requested_roll}: "
                    f"{result.academic_status.upper()} - {result.full_name} - Marks: {marks}"
                )
            elif result.fetch_status == "not_found":
                counters["not_found"] += 1
                print(f"[{completed}/{len(rolls)}] {result.requested_roll}: not found")
            else:
                counters["errors"] += 1
                print(
                    f"[{completed}/{len(rolls)}] {result.requested_roll}: "
                    f"{result.fetch_status} - {result.error}"
                )

        print("\nRequest summary")
        print(f"Students found:       {counters['found']}")
        print(f"Supply cases:         {counters['supply']}")
        print(f"Roll numbers missing: {counters['not_found']}")
        print(f"Errors:               {counters['errors']}")

        return results


def main() -> None:
    print("=" * 65)
    print("BISE Faisalabad Result Extractor - ResultDetails.aspx 2026")
    print("=" * 65)

    start_input = input("Starting Roll Number: ")
    end_input = input("Ending Roll Number:   ")

    try:
        start, end = validate_range(start_input, end_input)
    except ValueError as error:
        print(f"Error: {error}")
        return

    try:
        results = asyncio.run(run_scraper(start, end))
    except KeyboardInterrupt:
        print("\nCancelled by user.")
        return
    except Exception as error:
        print(f"Fatal error: {type(error).__name__}: {error}")
        return

    if not results:
        print("No request results were produced; no report files were created.")
        return

    excel_path = Path(f"bisefsd_results_{start}_{end}.xlsx")
    html_path = Path(f"bisefsd_results_{start}_{end}.html")

    try:
        save_to_excel(results, excel_path)
        save_to_html(results, html_path, start, end)
    except OSError as error:
        print(f"Could not save output files: {error}")
        return

    print("\nFiles created:")
    print(f"Excel: {excel_path.resolve()}")
    print(f"HTML:  {html_path.resolve()}")
    print("\nOpen the HTML file in Chrome, Edge, or Firefox.")


if __name__ == "__main__":
    main()