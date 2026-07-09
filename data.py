"""
Data Layer — file ingestion, column mapping, and fact extraction.
-----------------------------------------------------------------
Reads an Excel project plan, normalises column names, then computes
project-level facts (counts, percentages) that the agent layer uses
for RAG scoring. No LLM logic lives here.
"""

import os
import re

import pandas as pd

# ---------------------------------------------------------------------------
# Column mapping — small, per-file lookup, not a rewritten script.
# Only needed for columns whose NAME differs across files. Everything with
# the same name in both files doesn't need an entry.
# ---------------------------------------------------------------------------
# Only the columns the scoring logic actually reads need an entry here.
# Some fields list more than one CANDIDATE real column name, because some
# files have a duplicate/newer version of a field (e.g. UniSan has both
# "Variance" and a more-populated "Variance2"). When multiple candidates
# exist in a file, we pick whichever one actually has the least missing
# data — we don't hardcode which file uses which name.
COLUMN_CANDIDATES = {
    "task_name": ["Task Name"],
    "status": ["Status"],
    "pct_complete": ["% Complete"],
    "schedule_health": ["Schedule Health"],
    "variance": ["Variance", "Variance2"],
    "baseline_start": ["Baseline Start", "Baseline Start2", "Baseline Start Date"],
    "baseline_finish": ["Baseline Finish", "Baseline Finish2", "Baseline End Date"],
    "comments": ["Comments"],
    "critical": ["Critical ?"],
    "on_hold": ["On Hold?"],
    "not_applicable": ["Not Applicable?"],
    "end_date": ["End Date"],
    "rag_existing": ["RAG"],  # may not exist in every file — handled gracefully
}


class ProjectFileError(Exception):
    """Raised when a project plan file can't be read at all. Kept separate
    from normal Python errors so the caller can show a clean message
    instead of a raw traceback."""
    pass


def load_and_map(path: str) -> pd.DataFrame:
    """Load an Excel project plan and map only the columns we actually need.

    For fields with more than one candidate source column, pick whichever
    candidate present in this file has the LEAST missing data.

    Raises ProjectFileError with a clear message if the file is missing,
    unreadable, or not a valid Excel file — instead of letting a raw
    pandas/OS error crash the whole script.
    """
    if not os.path.exists(path):
        raise ProjectFileError(f"File not found: '{path}'. Check the path and try again.")

    try:
        df = pd.read_excel(path)
    except Exception as e:
        raise ProjectFileError(
            f"Could not read '{path}' as an Excel file. It may be corrupted, "
            f"password-protected, or not actually an .xlsx file. Original error: {e}"
        )

    if df.shape[0] == 0:
        raise ProjectFileError(f"'{path}' was read successfully but contains no rows.")

    for clean_name, candidates in COLUMN_CANDIDATES.items():
        present = [c for c in candidates if c in df.columns]

        if not present:
            df[clean_name] = None
            continue

        if len(present) == 1:
            best = present[0]
        else:
            # pick the candidate with the fewest missing values in this file
            best = min(present, key=lambda c: df[c].isna().mean())

        df[clean_name] = df[best]

    return df


# ---------------------------------------------------------------------------
# Turn rows into project-level facts. Pure counting, no judgment.
# ---------------------------------------------------------------------------
def _parse_variance_days(value):
    """Convert values like '-32d', '15d', '0' into an integer number of days."""
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return value
    match = re.search(r"-?\d+", str(value))
    return int(match.group()) if match else None


def compute_facts_loop(df: pd.DataFrame) -> dict:
    """Same job as compute_facts, but written as one explicit loop over rows
    so each step is easy to follow. No LLM involved anywhere in this function.
    """
    total_tasks = 0
    active_tasks = 0
    overdue_count = 0
    variance_values_seen = 0
    blocker_count = 0
    critical_blocked_count = 0
    red_health_count = 0
    yellow_health_count = 0
    health_values_seen = 0
    comments_seen = 0

    for _, row in df.iterrows():
        total_tasks += 1

        # Skip tasks explicitly marked Not Applicable — they don't count
        # toward project health at all.
        if row["not_applicable"] == 1:
            continue
        active_tasks += 1

        # --- Schedule slippage: look at this row's variance value ---
        variance_days = _parse_variance_days(row["variance"])
        if variance_days is not None:
            variance_values_seen += 1
            if variance_days < -10:
                overdue_count += 1

        # --- Blockers: a non-empty comment, or an On Hold status ---
        has_comment = pd.notna(row["comments"]) and str(row["comments"]).strip() != ""
        if has_comment:
            comments_seen += 1
            blocker_count += 1
        if row["status"] == "On Hold":
            blocker_count += 1

        # --- Critical task stuck on hold: the single worst signal ---
        if row["critical"] == 1 and row["status"] == "On Hold":
            critical_blocked_count += 1

        # --- Task-level Schedule Health, as a cross-check ---
        if pd.notna(row["schedule_health"]):
            health_values_seen += 1
            if row["schedule_health"] == "Red":
                red_health_count += 1
            elif row["schedule_health"] == "Yellow":
                yellow_health_count += 1

    # After the loop: turn raw counts into percentages, and flag anything
    # we didn't have enough real data to trust.
    missing_signals = []

    if active_tasks == 0:
        # Every row was marked Not Applicable, or something upstream went
        # wrong. There's nothing meaningful to score — say so plainly
        # instead of dividing by zero.
        return {
            "total_tasks": total_tasks,
            "active_tasks": 0,
            "pct_overdue": None,
            "overdue_count": 0,
            "blocker_count": 0,
            "critical_blocked_count": 0,
            "red_health_ratio": None,
            "yellow_health_ratio": None,
            "missing_signals": ["all_signals_no_active_tasks"],
        }

    if variance_values_seen == 0 or (variance_values_seen / active_tasks) < 0.1:
        missing_signals.append("schedule_variance")
        pct_overdue = None
    else:
        pct_overdue = round(100 * overdue_count / variance_values_seen, 1)

    if comments_seen == 0:
        missing_signals.append("blocker_comments")

    if health_values_seen == 0:
        missing_signals.append("schedule_health")
        red_health_ratio = None
        yellow_health_ratio = None
    else:
        red_health_ratio = round(red_health_count / health_values_seen, 2)
        yellow_health_ratio = round(yellow_health_count / health_values_seen, 2)

    return {
        "total_tasks": total_tasks,
        "active_tasks": active_tasks,
        "pct_overdue": pct_overdue,
        "overdue_count": overdue_count,
        "blocker_count": blocker_count,
        "critical_blocked_count": critical_blocked_count,
        "red_health_ratio": red_health_ratio,
        "yellow_health_ratio": yellow_health_ratio,
        "missing_signals": missing_signals,
    }
