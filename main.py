"""
Project Health RAG Reporting Agent
-----------------------------------
Reads a project plan (.xlsx), computes a Red/Amber/Green status using
deterministic rules, then generates a plain-English explanation of that
status. The color is always decided by rule-based Python — never by the
LLM. If OPENROUTER_API_KEY is set, the explanation is written by an LLM
via OpenRouter; otherwise it falls back to a template-based explanation
so the script always works, with or without an API key.

Usage:
    1. Create a file named ".env" in this same folder with one line:
           OPENROUTER_API_KEY=your_key_here
       (optional — without it, the script still works using a fallback
       template-based explanation, no LLM call at all)
    2. python main.py --file "S2P_Project.xlsx" --project-name "Titan"
    3. python main.py --file "Project_Plan_B.xlsx" --project-name "UniSan"
"""

import argparse
import json

from agent import run_agent
from data import ProjectFileError

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Project Health RAG Reporting Agent")
    parser.add_argument("--file", required=True, help="Path to the project plan .xlsx file")
    parser.add_argument("--project-name", required=True, help="Name of the project, e.g. Titan")
    parser.add_argument("--out", default=None, help="Optional path to save the report as JSON")
    args = parser.parse_args()

    try:
        report = run_agent(args.file, args.project_name)
    except ProjectFileError as e:
        print(f"Could not generate report: {e}")
        raise SystemExit(1)

    print(json.dumps(report, indent=2, default=str))

    if args.out:
        with open(args.out, "w") as f:
            json.dump(report, f, indent=2, default=str)
        print(f"\nSaved report to {args.out}")