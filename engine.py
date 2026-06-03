"""
Dot Camp 5 – Selection Engine
Exact rules from "Process de selection" sheet:

F1 Age:
  < 2 years  → SHORTLISTED baseline (continues to F3, F4 but stays ≤ Shortlisted)
  2–5 years  → SELECTED baseline → goes to F2
  5–7 years  → SHORTLISTED baseline (continues to F3, F4 but stays ≤ Shortlisted)
  > 7 years  → REJECTED immediately

F2 (2–5 years ONLY):
  Legally created = Yes → pass
  Legally created = No  → REJECTED

F3 (all non-rejected):
  At least one founder full-time AND resident in Tunisia → pass
  Otherwise → REJECTED

F4 Maturity (all non-rejected):
  Idea                          → REJECTED (for all age groups)
  POC finalized                 → SHORTLISTED (for all age groups)
  Functional MVP                → SELECTED if 2-5y / SHORTLISTED if <2y or 5-7y
  MVP currently being tested    → SELECTED if 2-5y / SHORTLISTED if <2y or 5-7y
  Go To Market (Early sales)    → SELECTED if 2-5y / SHORTLISTED if <2y or 5-7y
  Product/Service on the market → SELECTED if 2-5y / SHORTLISTED if <2y or 5-7y
  International Expansion       → SELECTED if 2-5y / SHORTLISTED if <2y or 5-7y

Key rule: <2y and 5-7y startups can only ever be Shortlisted or Rejected.
          Their baseline is Shortlisted and maturity cannot upgrade them to Selected.

Bonus points (add-ons, don't change the decision category):
  +1 each: employees, salaried, gender-mixed, clients, revenue, label, funding, programs
"""

from dataclasses import dataclass, field
from typing import List, Optional

# ── Option lists for UI dropdowns ─────────────────────────────────────────────
MATURITY_OPTIONS = [
    "Idea",
    "POC finalized",
    "Functional MVP",
    "MVP currently being tested",
    "Go To Market (Early sales)",
    "Product/Service on the market",
    "International Expansion",
]

AGE_OPTIONS = [
    "Less than 2 years",
    "2–5 years",
    "5–7 years",
    "More than 7 years",
]

EMPLOYEE_RANGES = ["None", "1–2", "3–5", "6–10", "+10"]
CLIENT_RANGES   = ["None", "1–2", "3–5", "6–10", "+10"]

# ── Maturity sets ──────────────────────────────────────────────────────────────
# These maturities → Selected for 2-5y, but only Shortlisted for <2y and 5-7y
_MATURITY_SELECTED_ELIGIBLE = {
    "Functional MVP",
    "MVP fonctionnel",
    "MVP currently being tested",
    "MVP en cours de test",
    "Go To Market (Early sales)",
    "Produit / Service commercialisé",
    "Product/Service on the market",
    "Produit / Service en expansion internationale",
    "International Expansion",
}

# These maturities → always Shortlisted (regardless of age group)
_MATURITY_SHORTLISTED = {
    "POC finalized",
    "POC réalisé",
}

# Age groups that are capped at Shortlisted (never Selected)
_SHORTLISTED_AGE_GROUPS = {"<2", "5-7"}

# ── Age normalisation ──────────────────────────────────────────────────────────
_AGE_GROUP = {
    "Less than 2 years":  "<2",
    "Less than  2 years": "<2",
    "Moins de 2 ans":     "<2",
    "2–5 years":          "2-5",
    "2-5 years":          "2-5",
    "2–5 ans":            "2-5",
    "2-5 ans":            "2-5",
    "5–7 years":          "5-7",
    "5-7 years":          "5-7",
    "5–7 ans":            "5-7",
    "5-7 ans":            "5-7",
    "More than 7 years":  ">7",
    "Plus de 7 ans":      ">7",
}


@dataclass
class EvaluationResult:
    startup_name:     str
    final_decision:   str          # "Selected ★" | "Selected" | "Shortlisted ✓" | "Shortlisted" | "Rejected"
    emoji:            str
    bonus_score:      int
    bonus_details:    List[str]
    rejection_reason: Optional[str]
    filter1_result:   str
    filter2_result:   str
    filter3_result:   str
    filter4_result:   str


def _truthy(v) -> bool:
    """Return True for Yes/Oui/1/True values."""
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    return str(v).strip().lower() in {"yes", "oui", "true", "1", "1.0", "vrai"}


def _has_clients(v) -> bool:
    s = str(v).strip().lower()
    return s not in {"none", "aucun", "0", "nan", "", "null"}


def evaluate(app: dict) -> EvaluationResult:
    name    = app.get("startup_name", "Unknown")
    age_raw = str(app.get("age", "")).strip()
    age_grp = _AGE_GROUP.get(age_raw, age_raw)

    # ── F1: Age ────────────────────────────────────────────────────────────────
    if age_grp == ">7":
        return EvaluationResult(
            startup_name=name,
            final_decision="Rejected",
            emoji="❌",
            bonus_score=0,
            bonus_details=[],
            rejection_reason="Startup is more than 7 years old",
            filter1_result="❌ Rejected (>7 years)",
            filter2_result="—",
            filter3_result="—",
            filter4_result="—",
        )

    # Determine the age-based ceiling
    # <2y and 5-7y are capped at Shortlisted; 2-5y can reach Selected
    is_shortlisted_baseline = age_grp in _SHORTLISTED_AGE_GROUPS

    if is_shortlisted_baseline:
        f1 = f"⚠️ Shortlisted baseline ({age_raw}) — max outcome: Shortlisted"
    else:
        f1 = "✅ Pass (2–5 years, Selected baseline)"

    # ── F2: Legal check — 2–5 years ONLY ──────────────────────────────────────
    legally = _truthy(app.get("legally_created", False))

    if age_grp == "2-5":
        if not legally:
            return EvaluationResult(
                startup_name=name,
                final_decision="Rejected",
                emoji="❌",
                bonus_score=0,
                bonus_details=[],
                rejection_reason="2–5 year startup is not legally established",
                filter1_result=f1,
                filter2_result="❌ Rejected (not legally created)",
                filter3_result="—",
                filter4_result="—",
            )
        f2 = "✅ Pass (legally created)"
    else:
        # <2 years and 5–7 years: legal status is NOT a rejection criterion
        f2 = "➖ N/A (legal check applies to 2–5y only)"

    # ── F3: Founder full-time AND resident in Tunisia ──────────────────────────
    fulltime = _truthy(app.get("full_time_founder", False))
    resident = _truthy(app.get("founder_in_tunisia", False))

    if not fulltime or not resident:
        reason = []
        if not fulltime: reason.append("no full-time founder")
        if not resident: reason.append("no founder resident in Tunisia")
        return EvaluationResult(
            startup_name=name,
            final_decision="Rejected",
            emoji="❌",
            bonus_score=0,
            bonus_details=[],
            rejection_reason=f"Founder filter failed: {', '.join(reason)}",
            filter1_result=f1,
            filter2_result=f2,
            filter3_result=f"❌ Rejected ({', '.join(reason)})",
            filter4_result="—",
        )

    f3 = "✅ Pass"

    # ── F4: Solution maturity ──────────────────────────────────────────────────
    maturity = str(app.get("maturity", "")).strip()

    if maturity in _MATURITY_SELECTED_ELIGIBLE:
        if is_shortlisted_baseline:
            # <2y and 5-7y: good maturity keeps them Shortlisted (cannot upgrade to Selected)
            base_decision = "Shortlisted"
            f4 = f"⚠️ Shortlisted ({maturity}) — age group capped at Shortlisted"
        else:
            # 2-5y: good maturity → Selected
            base_decision = "Selected"
            f4 = f"✅ Selected ({maturity})"
    elif maturity in _MATURITY_SHORTLISTED:
        # POC → always Shortlisted regardless of age
        base_decision = "Shortlisted"
        f4 = f"⚠️ Shortlisted ({maturity})"
    else:
        # Idea or unrecognised → Rejected for all age groups
        return EvaluationResult(
            startup_name=name,
            final_decision="Rejected",
            emoji="❌",
            bonus_score=0,
            bonus_details=[],
            rejection_reason=f"Maturity too low: '{maturity}'",
            filter1_result=f1,
            filter2_result=f2,
            filter3_result=f3,
            filter4_result=f"❌ Rejected (maturity: {maturity})",
        )

    # ── Bonus scoring ──────────────────────────────────────────────────────────
    bonus       = 0
    bonus_items = []

    def _add(condition, label):
        nonlocal bonus
        if condition:
            bonus += 1
            bonus_items.append(label)

    emp = str(app.get("total_employees", "None")).strip().lower()
    sal = str(app.get("salaried_employees", "None")).strip().lower()

    _add(emp not in {"none", "aucun", "0", "nan", "", "null"},  "Has employees")
    _add(sal not in {"none", "aucun", "0", "nan", "", "null"},  "Has salaried staff")
    _add(_truthy(app.get("gender_mixed", False)),               "Gender-mixed team")
    _add(_has_clients(app.get("num_clients", "None")),          "Has clients")
    _add(_truthy(app.get("generating_revenue", False)),         "Generating revenue")
    _add(_truthy(app.get("startup_label", False)),              "Startup Act label")
    _add(_truthy(app.get("raised_funding", False)),             "Raised funding")
    _add(_truthy(app.get("participated_programs", False)),      "Participated in programs")

    # ── Final decision with bonus star ────────────────────────────────────────
    if base_decision == "Selected":
        # Only 2-5y startups can reach this branch
        final = "Selected ★" if bonus >= 3 else "Selected"
        emoji = "⭐" if bonus >= 3 else "✅"
    else:
        # Shortlisted (all <2y, all 5-7y, and POC-maturity 2-5y startups)
        final = "Shortlisted ✓" if bonus >= 2 else "Shortlisted"
        emoji = "🌟" if bonus >= 2 else "⚠️"

    return EvaluationResult(
        startup_name=name,
        final_decision=final,
        emoji=emoji,
        bonus_score=bonus,
        bonus_details=bonus_items,
        rejection_reason=None,
        filter1_result=f1,
        filter2_result=f2,
        filter3_result=f3,
        filter4_result=f4,
    )