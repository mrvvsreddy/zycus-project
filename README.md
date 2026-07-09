# Project Health Reporting Agent

An AI-powered system that reads project plans, computes RAG (Red/Amber/Green) health status using deterministic rules, generates plain-English explanations, and synthesizes portfolio-level insights into an executive PowerPoint presentation.

---

## Architecture

The system is organized into three phases matching the assignment structure:

```
Phase 1 (Framework)        Phase 2 (Agent)              Phase 3 (Presentation)
─────────────────          ────────────────              ──────────────────────
RAGMethodology.md     ┌──► data.py                      generate_presentation.py
                      │    ├─ load_and_map()             ├─ load_reports()
                      │    └─ compute_facts_loop()       ├─ aggregate_data()
                      │          │                       ├─ generate_insights_with_llm()
                      │          ▼                       └─ create_presentation()
                      │    agent.py                            │
                      │    ├─ score_rag()                      ▼
                      │    ├─ explain_with_llm()          pptx_tools.py
                      │    └─ build_report()              ├─ chart_donut / hbar / vbar
                      │          │                        ├─ add_metric_card / table
                      │          ▼                        └─ embed_chart
                      │    main.py (CLI)
                      │          │
                      │          ▼
Excel files ──────────┘    reports/*.json ──────────────► presentations/*.pptx
```

### File Roles

| File | Purpose |
|---|---|
| `data.py` | Reads Excel project plans, maps varying column names to canonical fields, extracts project-level facts (counts, percentages). No LLM logic. |
| `agent.py` | Applies deterministic RAG scoring rules against the facts. Calls an LLM via OpenRouter for a plain-English explanation (falls back to templates if no API key). Assembles the final JSON report. |
| `main.py` | CLI entry point for Phase 2. Runs the agent against a single Excel file and outputs a JSON report. |
| `generate_presentation.py` | Phase 3 orchestrator. Loads all JSON reports from `reports/`, aggregates portfolio metrics, calls the LLM for cross-project insights, and builds an 8-slide dark-themed PPTX. |
| `pptx_tools.py` | Reusable slide-building toolkit — design tokens, primitive shapes, composite cards, and matplotlib chart generators with dynamic sizing. |
| `RAGMethodology.md` | One-page RAG framework defining the scoring signals, thresholds, and assumptions (Phase 1 deliverable). |

---

## How to Run

### Prerequisites

- Python ≥ 3.11
- [`uv`](https://docs.astral.sh/uv/) package manager (handles virtualenv and dependencies automatically)

### Setup

```bash
# Clone and enter the project
cd zycus-project

# (Optional) Create a .env file for LLM-powered explanations
echo "OPENROUTER_API_KEY=your_key_here" > .env
```

> **Note:** The agent works without an API key — it falls back to template-based explanations. The LLM is used only for writing natural-language summaries; the RAG color is always computed deterministically.

### Phase 2 — Run the Agent (Weekly Reports)

Process each project plan individually and save the JSON report:

```bash
# Titan (S2P_Project.xlsx)
uv run main.py --file "S2P_Project.xlsx" --project-name "Titan" --out reports/Titan_$(date +%F).json

# UniSan (Project Plan B.xlsx)
uv run main.py --file "Project Plan B.xlsx" --project-name "UniSan" --out reports/UniSan_$(date +%F).json
```

Each command outputs a structured JSON report containing:
- The computed RAG status (Red / Amber / Green)
- Supporting facts (task counts, overdue %, blocker counts, health ratios)
- A plain-English explanation of why the project received that status
- A list of any signals that could not be assessed due to missing data

### Phase 3 — Generate the Monthly Presentation

Aggregate all reports and produce the executive slide deck:

```bash
uv run generate_presentation.py
```

This reads every `*.json` file in `reports/`, aggregates portfolio-level metrics, calls the LLM for cross-project insights and recommendations, and outputs a PPTX to `presentations/`.

### Scheduling (Bonus)

To run the agent on a weekly schedule, add a cron job:

```bash
# Every Monday at 9:00 AM
0 9 * * 1 cd /path/to/zycus-project && uv run main.py --file "S2P_Project.xlsx" --project-name "Titan" --out reports/Titan_$(date +\%F).json
0 9 * * 1 cd /path/to/zycus-project && uv run main.py --file "Project Plan B.xlsx" --project-name "UniSan" --out reports/UniSan_$(date +\%F).json
```

---

## Design Decisions

### RAG Scoring: Deterministic Rules, Not LLM Guessing

The RAG color is computed by strict, auditable business rules in `agent.py → score_rag()`. The LLM is only called *after* the color is decided, to write a human-readable explanation. This ensures:
- Reproducible results — same data always produces the same color
- No hallucinated metrics — the LLM never sees raw rows or decides severity
- Graceful degradation — if the LLM is unavailable, a template-based fallback fires automatically

### Robust Column Mapping

Real project plans use inconsistent column names (e.g., `Variance` vs `Variance2`, `Baseline Start` vs `Baseline Start Date`). Instead of hardcoding file-specific mappings, `data.py` defines candidate columns for each logical field and picks whichever candidate has the least missing data in that particular file.

### Missing Data Is Never "Healthy"

If a signal (e.g., schedule variance) has insufficient data, it is excluded from scoring entirely and disclosed in the report's `missing_signals` array. The system never silently assumes missing data means everything is fine.

### Modular Presentation Toolkit

Rather than writing monolithic slide-generation code, `pptx_tools.py` exports composable primitives — `add_card`, `add_metric_card`, `add_slide_header`, `chart_donut`, `chart_hbar`, etc. These are assembled like building blocks in `generate_presentation.py`, making it easy to add new slide types or rearrange layouts.

### Dynamic Chart Sizing

Chart functions accept a `figsize` parameter calculated from the available slide real estate (slide dimensions minus margins and headers). The matplotlib figure is rendered at exactly the target aspect ratio, then embedded with matching width and height. This eliminates the distortion that occurs when PowerPoint stretches a chart image into a different aspect ratio.

---

## Sample Outputs

### Weekly Reports (`reports/`)

| Project | RAG Status | Key Facts |
|---|---|---|
| **Titan** | 🔴 Red | 41.7% overdue rate (110 of 484 active tasks), 5 blockers |
| **UniSan** | 🟢 Green | 0% overdue, 0 blockers, blocker comments data unavailable |

### Monthly Presentation (`presentations/`)

An 8-slide dark-themed executive deck:

| Slide | Content |
|---|---|
| 1 | Title — Monthly Portfolio Synthesis |
| 2 | Executive Summary — KPI cards + RAG donut chart |
| 3 | Portfolio Trends: Overdue Tasks — horizontal bar chart |
| 4 | Portfolio Trends: Blockers — grouped vertical bar chart |
| 5 | Emerging Risks — color-coded data table |
| 6 | Executive Insights — LLM-generated trend analysis |
| 7 | Recommendations — LLM-generated action items |
| 8 | Closing — Thank You |
