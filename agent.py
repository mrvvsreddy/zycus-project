"""
Agent Layer — RAG scoring, LLM explanation, and report assembly.
-----------------------------------------------------------------
Takes the project-level facts produced by data.py and applies
rule-based scoring to decide Red/Amber/Green. Optionally calls
OpenRouter for an LLM-written explanation; falls back to a
template if no API key is available or the call fails.
"""

import json
import os
from datetime import datetime

import requests
from dotenv import load_dotenv

from data import load_and_map, compute_facts_loop

load_dotenv()  # reads a .env file in the current directory, if one exists

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
OPENROUTER_MODEL = os.environ.get("OPENROUTER_MODEL")  # change to any model OpenRouter supports


def score_rag(facts: dict) -> str:
    """Apply the Phase 1 rules to decide Red / Amber / Green.
    Pure rule-based logic — no LLM, same result every time for the same facts.
    """
    pct_overdue = facts.get("pct_overdue")
    red_ratio = facts.get("red_health_ratio")
    yellow_ratio = facts.get("yellow_health_ratio")
    critical_blocked = facts.get("critical_blocked_count", 0)
    blockers = facts.get("blocker_count", 0)

    # RED: any single severe issue is enough on its own
    if critical_blocked > 0:
        return "Red"
    if pct_overdue is not None and pct_overdue > 15:
        return "Red"
    if red_ratio is not None and red_ratio > 0.5:
        return "Red"

    # AMBER: smaller issues, nothing severe
    if pct_overdue is not None and pct_overdue >= 5:
        return "Amber"
    if blockers > 0:
        return "Amber"
    if yellow_ratio is not None and yellow_ratio > 0.3:
        return "Amber"

    return "Green"


# ---------------------------------------------------------------------------
# Fallback explanation — plain templates, zero external calls.
# Used automatically if no OPENROUTER_API_KEY is set, or if the API call
# fails for any reason. The script must never break just because the LLM
# step is unavailable.
# ---------------------------------------------------------------------------
def explain_result_fallback(project_name: str, rag_status: str, facts: dict) -> str:
    sentences = [f"{project_name} is currently {rag_status}."]

    if facts.get("critical_blocked_count", 0) > 0:
        sentences.append(
            f"{facts['critical_blocked_count']} critical task(s) are on hold, "
            f"which is the main driver of this status."
        )

    if facts.get("pct_overdue") is not None:
        sentences.append(
            f"{facts['pct_overdue']}% of tasks with schedule data are overdue "
            f"by more than 10 days ({facts.get('overdue_count', 0)} tasks)."
        )

    if facts.get("blocker_count", 0) > 0:
        sentences.append(
            f"{facts['blocker_count']} task(s) show open blockers or comments "
            f"indicating a delay."
        )

    if facts.get("red_health_ratio") is not None:
        sentences.append(
            f"{int(facts['red_health_ratio'] * 100)}% of individually-tracked "
            f"tasks are already flagged Red at the task level, "
            f"which is roughly consistent with this overall status."
        )

    missing = facts.get("missing_signals", [])
    if missing:
        sentences.append(
            f"Note: {', '.join(missing)} could not be reliably assessed for this "
            f"project due to insufficient data, so this status carries slightly "
            f"less evidence than a project with complete data."
        )

    return " ".join(sentences)


# ---------------------------------------------------------------------------
# LLM explanation via OpenRouter. The LLM only writes the sentence —
# it never sees raw rows and never decides the color. If this fails for any
# reason (no key, network issue, bad response), we fall back automatically.
# ---------------------------------------------------------------------------
def explain_with_llm(project_name: str, rag_status: str, facts: dict) -> str:
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        return explain_result_fallback(project_name, rag_status, facts)

    missing = facts.get("missing_signals", [])
    missing_note = (
        f"Note: the following signals were unavailable for this project and "
        f"were excluded from scoring: {', '.join(missing)}."
        if missing else ""
    )

    prompt = f"""You are writing a short project health summary for a VP audience.

Project: {project_name}
RAG status: {rag_status}
Supporting facts: {json.dumps({k: v for k, v in facts.items() if k != "missing_signals"}, default=str)}
{missing_note}

Write 2-3 plain-English sentences explaining why this project has this RAG
status. Reference the specific facts above. Write it as a human PM update,
not as a description of rules or an algorithm. If signals were missing,
mention that plainly and note the status has slightly less evidence behind
it as a result."""

    try:
        response = requests.post(
            OPENROUTER_URL,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": OPENROUTER_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": 300,
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        # Never let an API problem break the whole report — fall back instead.
        fallback = explain_result_fallback(project_name, rag_status, facts)
        return f"{fallback} [LLM explanation unavailable: {e}]"


# ---------------------------------------------------------------------------
# Build and return/save the final report.
# ---------------------------------------------------------------------------
def build_report(project_name: str, rag_status: str, facts: dict, explanation: str) -> dict:
    return {
        "project_name": project_name,
        "generated_at": datetime.now().isoformat(),
        "rag_status": rag_status,
        "facts": facts,
        "reasoning": explanation,
    }


def run_agent(file_path: str, project_name: str) -> dict:
    df = load_and_map(file_path)
    facts = compute_facts_loop(df)
    rag_status = score_rag(facts)
    explanation = explain_with_llm(project_name, rag_status, facts)
    return build_report(project_name, rag_status, facts, explanation)
