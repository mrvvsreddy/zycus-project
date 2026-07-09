# Project Health RAG Methodology

## Overview

This document defines how each project's Red / Amber / Green status is determined. The RAG color is always computed by deterministic rules in code (`agent.py → score_rag()`). The LLM's only role is to write a plain-English explanation *after* the color has already been decided — it never influences the scoring itself.

## Scoring Signals

| Signal | Source Columns | Red Trigger | Amber Trigger | If Data Is Missing |
|---|---|---|---|---|
| **Schedule slippage** | `Variance` / `Variance2` parsed to days | > 15 % of tasks overdue by 10+ days | 5–15 % of tasks overdue by 10+ days | Excluded from scoring, flagged as `schedule_variance` |
| **Blockers** | `Comments` (non-empty text), `Status = "On Hold"` | — (blockers alone don't trigger Red) | Any blocker count > 0 | If no comments exist at all, flagged as `blocker_comments` |
| **Critical tasks blocked** | `Critical ? = 1` AND `Status = "On Hold"` | Any critical task on hold (count > 0) | — | Counted directly; no missing-data case |
| **Task-level health** | `Schedule Health` column (Red / Yellow / Green) | > 50 % of health-tracked tasks are Red | > 30 % are Yellow | Excluded from scoring, flagged as `schedule_health` |
| **Budget burn** | Not available in sample data | Not scored | Not scored | Defined in framework but excluded — never faked |
| **Stakeholder sentiment** | No dedicated field; partially captured via `Comments` | — folded into the blocker signal | — | No signal, no penalty |

## Decision Logic

Every project starts at **Green** and is promoted upward through the following cascade:

### → Red (any single condition is sufficient)
1. At least one critical task is on hold (`critical_blocked_count > 0`)
2. More than 15 % of variance-tracked tasks are overdue by 10+ days (`pct_overdue > 15`)
3. More than 50 % of health-tracked tasks are individually flagged Red (`red_health_ratio > 0.5`)

### → Amber (if none of the Red conditions fired)
1. 5–15 % of variance-tracked tasks are overdue (`pct_overdue >= 5`)
2. Any non-zero blocker count (`blocker_count > 0`)
3. More than 30 % of health-tracked tasks are Yellow (`yellow_health_ratio > 0.3`)

### → Green
If none of the above conditions are met.

## Handling Incomplete Data

- **Missing data is never treated as "healthy."** If fewer than 10 % of active tasks have a variance value, the schedule slippage signal is excluded from scoring entirely and disclosed in the report as a `missing_signal`.
- Rows marked `Not Applicable? = 1` are excluded from all counts — they don't inflate or deflate the health picture.
- Every report's `missing_signals` array explicitly lists which signals could not be assessed, so the reader knows how much evidence backs the final status.

## Assumptions

- The two sample project plans (`S2P_Project.xlsx` / Titan and `Project Plan B.xlsx` / UniSan) use slightly different column names for the same data. The agent handles this via a candidate-mapping system that picks whichever source column has the least missing data, rather than hardcoding file-specific mappings.
- A task is considered "overdue" when its `Variance` value is worse than −10 days. Minor slippage under 10 days is not counted.
- Non-empty `Comments` text is treated as a blocker indicator. This is a conservative assumption — it may slightly overcount blockers, but it ensures no flagged issues are silently ignored.

## In Short

Green by default. One severe issue alone makes it Red. A few smaller issues make it Amber. Whatever the data can't tell us gets flagged as unknown, never quietly assumed to be fine.