"""
build_dashboard.py
-------------------
Single-input automated build: reads ONE Excel file (Dashboard_Input_Template.xlsx,
sheet "Paper_Topic_Assignments") and produces dashboard_data.json + dashboard.html.

WHAT THIS SCRIPT DOES vs DOES NOT DO
  - Computes corpus stats live from the input sheet: totals, pre/post-ChatGPT split,
    category breakdown, source-DB breakdown, per-topic counts and splits.
  - Does NOT recompute evidence scores (D1/D2/D3, composite, tier). Those are locked
    constants verified from Stage 6 Cell 6 output on 20 June 2026 and are pasted
    below unchanged. This is a deliberate decision, not an oversight: the figures
    are already cited in the dissertation Discussion section, and re-deriving them
    from a generic pipeline would reintroduce the exact non-determinism risk that
    BERTopic Stage 2 was locked to avoid.
  - Outcome_Reported in the input sheet is used ONLY for a sanity cross-check
    against the locked D1 percentage (see d1_cross_check in the output JSON).
    It never overwrites the locked value.

RUN
    cd RCSI_Dashboard_Pipeline
    python scripts/build_dashboard.py
    -> output/dashboard_data.json
    -> output/dashboard.html   (single self-contained file, no server needed)

INPUT CONTRACT
    input/Dashboard_Input_Template.xlsx, sheet "Paper_Topic_Assignments", columns:
    Title, Year, Period, Classification, Topic_Number, Topic_Label,
    Topic_Confidence, Has_Abstract, Outcome_Reported, Source_DB
    Full column spec: see README_Column_Spec sheet in the same workbook.

LOCKED TOPIC STRUCTURE (never re-derive from text/clustering in this script):
  n=116  Topic 1  AI Benchmarking & Exam Performance
  n=111  Topic 0  Generative AI Adoption & Student Perceptions
  n= 97  Topic 2  AI in Medical Imaging & Surgical Education
  n= 96  Topic 4  Simulation, Virtual Reality & Clinical Training
  n= 56  Topic 5  Perceptions, Readiness & Knowledge-Attitude Studies
  n= 31  Topic 3  Pharmacy Education & Academic Integrity
"""

import json
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT          = Path(__file__).resolve().parent.parent
INPUT_FILE    = ROOT / "input" / "Dashboard_Input_Template.xlsx"
TEMPLATE_FILE = ROOT / "templates" / "dashboard_template.html"
OUTPUT_JSON   = ROOT / "output" / "dashboard_data.json"
OUTPUT_HTML   = ROOT / "output" / "dashboard.html"

# ── LOCKED TOPIC DEFINITIONS (unchanged from extract_to_json.py) ─────────────
LOCKED_TOPICS = {
    1: {"label": "AI Benchmarking & Exam Performance", "count": 116,
        "keywords": ["examination", "mcqs", "exam", "llms", "question",
                      "licensing", "multiplechoice", "chat", "bard", "answer"]},
    0: {"label": "Generative AI Adoption & Student Perceptions", "count": 111,
        "keywords": ["llms", "chatbots", "chat", "chatgpts", "questionnaire",
                      "tutor", "using chatgpt", "respondents", "chatbot", "simulations"]},
    2: {"label": "AI in Medical Imaging & Surgical Education", "count": 97,
        "keywords": ["images", "anatomy", "physicians", "systematic", "surgical",
                      "image", "anatomical", "synthetic", "scoping", "publications"]},
    4: {"label": "Simulation, Virtual Reality & Clinical Training", "count": 96,
        "keywords": ["simulation", "generative artificial", "platform", "llms",
                      "reality", "dental", "simulated", "statements",
                      "virtual reality", "virtual patient"]},
    5: {"label": "Perceptions, Readiness & Knowledge-Attitude Studies", "count": 56,
        "keywords": ["readiness", "questionnaire", "undergraduate medical", "radiology",
                      "knowledge attitudes", "intelligence medicine", "saudi",
                      "doctors", "medical curricula", "respondents"]},
    3: {"label": "Pharmacy Education & Academic Integrity", "count": 31,
        "keywords": ["pharmacy", "pharmacy education", "pharmacy students", "osce",
                      "documents", "articles", "osces", "integrity",
                      "academic integrity", "questionnaire"]},
}

# ── LOCKED EVIDENCE SCORES (verified Stage 6 Cell 6, 20 June 2026 — DO NOT DERIVE) ──
LOCKED_EVIDENCE = {
    1: {"d1_pct": 94.0, "d1_rating": "HIGH", "d2_pct": 7.8, "d2_rating": "LOW",
        "d3_rating": "HIGH", "composite": 2.4, "tier": "MEDIUM",
        "note": "D1/D2 inversion. Highest D1 in corpus (94%). Lowest D2 aside from Topic 5 (7.8%). "
                "Replicated across 3 BERTopic re-fits. Headline dissertation finding."},
    0: {"d1_pct": 64.9, "d1_rating": "MEDIUM", "d2_pct": 13.5, "d2_rating": "MEDIUM",
        "d3_rating": "MEDIUM", "composite": 2.0, "tier": "MEDIUM",
        "note": "All three dimensions MEDIUM. Consistent middle-tier evidence."},
    2: {"d1_pct": 50.5, "d1_rating": "LOW", "d2_pct": 29.9, "d2_rating": "HIGH",
        "d3_rating": "MEDIUM", "composite": 1.8, "tier": "MEDIUM",
        "note": "REVERSE D1/D2 inversion. Highest D2 in corpus (29.9%) but lowest D1 alongside "
                "Simulation (50.5%). Good methodology, poor outcome reporting."},
    4: {"d1_pct": 50.0, "d1_rating": "LOW", "d2_pct": 15.6, "d2_rating": "MEDIUM",
        "d3_rating": "MEDIUM", "composite": 1.5, "tier": "LOW",
        "note": "Joint-lowest composite (1.5 LOW) with Pharmacy. Half the papers report no outcome."},
    5: {"d1_pct": 82.1, "d1_rating": "HIGH", "d2_pct": 3.6, "d2_rating": "LOW",
        "d3_rating": "LOW", "composite": 2.0, "tier": "MEDIUM",
        "note": "Third D1/D2 inversion. D2=3.6% is the lowest in the entire corpus."},
    3: {"d1_pct": 58.1, "d1_rating": "LOW", "d2_pct": 16.1, "d2_rating": "MEDIUM",
        "d3_rating": "MEDIUM", "composite": 1.5, "tier": "LOW",
        "note": "Joint-lowest composite (1.5 LOW) with Simulation. Smallest topic (n=31)."},
}

LOCKED_GAP_ANALYSIS = {
    "study_design_distribution": {
        "Survey / Questionnaire": 134, "Observational Study": 89,
        "Case Study / Case Series": 72, "Comparative Study": 61,
        "Mixed Methods": 48, "Simulation-Based Study": 38,
        "Review / Meta-Analysis": 29, "Expert Opinion / Commentary": 17,
        "RCT / Randomized Trial": 13, "Other": 6,
    },
    "headline_gaps": [
        {"rank": 1, "gap": "Topic 1 (Benchmarking) D1/D2 inversion", "affected_topics": [1]},
        {"rank": 2, "gap": "Topic 5 (Readiness) lowest D2 in corpus", "affected_topics": [5]},
        {"rank": 3, "gap": "Simulation: largest volume-vs-evidence gap", "affected_topics": [4]},
        {"rank": 4, "gap": "Survey dominance limits causal inference", "affected_topics": [0, 5]},
        {"rank": 5, "gap": "Imaging: good design, poor outcome reporting", "affected_topics": [2]},
    ],
}


def load_input():
    df = pd.read_excel(INPUT_FILE, sheet_name="Paper_Topic_Assignments")
    if "Year" not in df.columns:
        # A flag/banner row sits above the real header row in this sheet — re-read
        # skipping it.
        df = pd.read_excel(INPUT_FILE, sheet_name="Paper_Topic_Assignments", header=1)
    df = df[df["Year"].notna()].copy()

    # Do NOT trust the raw Topic_Number column as-is. Different exports of this
    # sheet have used different numbering (0-5 in one export, 1-6 in another) for
    # the same six clusters. The Topic_Label text is the stable identifier, so
    # remap to the locked 0-5 IDs by matching label, and fail loudly if a label
    # doesn't match any locked label (rather than silently mis-assigning papers).
    label_to_id = {t["label"]: tid for tid, t in LOCKED_TOPICS.items()}
    unknown = sorted(set(df["Topic_Label"]) - set(label_to_id))
    if unknown:
        raise ValueError(
            f"Topic_Label values not found in the locked label set: {unknown}. "
            "Fix the input file's labels before running — do not guess a mapping."
        )
    df["Topic_Number"] = df["Topic_Label"].map(label_to_id).astype(int)
    return df


def check_locked(df):
    dist = df["Topic_Number"].value_counts().sort_index().to_dict()
    locked_dist = {tid: t["count"] for tid, t in LOCKED_TOPICS.items()}
    return dist == locked_dist, dist, locked_dist


def build_corpus_overview(df):
    total = len(df)
    pre  = int((df["Period"].astype(str).str.startswith("Pre")).sum())
    post = total - pre
    return {
        "total_papers": total,
        "pre_chatgpt": pre,
        "post_chatgpt": post,
        "pct_post_chatgpt": round(post / total * 100, 1) if total else 0,
        "chatgpt_cutoff": "November 2022",
        "by_category": df["Classification"].value_counts().to_dict(),
        "by_year": {int(k): int(v) for k, v in df["Year"].value_counts().sort_index().items()},
        "by_source_db": df["Source_DB"].value_counts().to_dict(),
    }


def build_topics(df):
    topics = []
    for tid, t in LOCKED_TOPICS.items():
        sub = df[df["Topic_Number"] == tid]
        ev = LOCKED_EVIDENCE[tid]
        pre = int((sub["Period"].astype(str).str.startswith("Pre")).sum()) if len(sub) else None
        post = (len(sub) - pre) if pre is not None else None

        d1_excel = None
        if "Outcome_Reported" in sub.columns and len(sub) > 0:
            d1_excel = round((sub["Outcome_Reported"].astype(str) == "Yes").sum() / len(sub) * 100, 1)
        d1_mismatch = bool(d1_excel is not None and abs(d1_excel - ev["d1_pct"]) > 1.0)

        topics.append({
            "id": tid, "label": t["label"], "count_locked": t["count"],
            "count_input_file": int(len(sub)),
            "pre_chatgpt": pre, "post_chatgpt": post,
            "by_category": sub["Classification"].value_counts().to_dict() if len(sub) else {},
            "keywords": t["keywords"],
            "evidence": {
                "d1_pct": ev["d1_pct"], "d1_rating": ev["d1_rating"],
                "d2_pct": ev["d2_pct"], "d2_rating": ev["d2_rating"],
                "d3_rating": ev["d3_rating"], "composite": ev["composite"], "tier": ev["tier"],
                "note": ev["note"], "source": "Stage 6 Cell 6 — verified 20 June 2026 — LOCKED",
                "d1_cross_check": {"computed_from_input_file": d1_excel,
                                    "stage6_locked_value": ev["d1_pct"],
                                    "mismatch_flag": d1_mismatch},
            },
        })
    return sorted(topics, key=lambda x: -x["count_locked"])


def build_evidence_summary(topics):
    tiers = {"HIGH": 0, "MEDIUM": 0, "LOW": 0}
    for t in topics:
        tiers[t["evidence"]["tier"]] += 1
    return {
        "tier_distribution": tiers, "all_verified": True,
        "verification_date": "20 June 2026",
        "verification_source": "Stage 6 notebook Cell 6 (LOCKED, not recomputed by this script)",
    }


def main():
    print("Loading input file...")
    df = load_input()
    print(f"  Rows loaded: {len(df)}")

    is_locked, dist, locked_dist = check_locked(df)
    if is_locked:
        print("  Topic distribution: LOCKED structure confirmed (matches 116/111/97/96/56/31)")
    else:
        print("  *** WARNING: input file topic distribution does NOT match the locked structure ***")
        print(f"      Input file:  {dist}")
        print(f"      Locked:      {locked_dist}")
        print("      Corpus stats below (counts, pre/post splits, category breakdowns) are computed")
        print("      from the INPUT FILE and will be wrong until you re-export Paper_Topic_Assignments")
        print("      from the saved locked model. Evidence scores (D1/D2/D3/tier) are unaffected —")
        print("      they are locked constants, not derived from this file.")

    corpus_overview = build_corpus_overview(df)
    topics = build_topics(df)
    evidence_summary = build_evidence_summary(topics)

    payload = {
        "meta": {
            "generated_at": datetime.now().isoformat(),
            "source_file": INPUT_FILE.name,
            "input_topic_distribution_locked": is_locked,
            "locked_structure_target": "116/111/97/96/56/31",
            "evidence_scores_are_locked_constants": True,
        },
        "corpus_overview": corpus_overview,
        "topics": topics,
        "evidence_summary": evidence_summary,
        "gap_analysis": LOCKED_GAP_ANALYSIS,
    }

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUTPUT_JSON}")

    if not TEMPLATE_FILE.exists():
        print(f"Template not found at {TEMPLATE_FILE} — skipping HTML build.")
        return

    html = TEMPLATE_FILE.read_text(encoding="utf-8")
    html = html.replace("__DASHBOARD_DATA_JSON__", json.dumps(payload, ensure_ascii=False))
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"Wrote {OUTPUT_HTML}")
    if not is_locked:
        print("\nReminder: dashboard.html was built from a NON-LOCKED input file. "
              "Fix the input and re-run before sharing this build.")


if __name__ == "__main__":
    main()
