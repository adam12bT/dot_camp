"""
Mock data generator for Dot Camp 5 selection pipeline.
Produces realistic startup applications that simulate:
  - Google Form CSV exports
  - Typeform JSON exports
  - Direct dict submissions
"""

import random
import json
import csv
import io
from datetime import datetime, timedelta

from engine import AGE_OPTIONS, MATURITY_OPTIONS, EMPLOYEE_RANGES, CLIENT_RANGES

random.seed(42)

SECTORS = [
    "AI/Data", "FinTech", "HealthTech", "EdTech", "GreenTech",
    "Agri/FoodTech", "HR Tech", "DeepTech", "SaaS", "TravelTech",
    "PropTech", "Industry 4.0", "ICC", "Transport Tech",
]

GOVERNORATES = [
    "Tunis", "Ariana", "Ben Arous", "Manouba", "Nabeul",
    "Bizerte", "Sousse", "Monastir", "Sfax", "Kairouan",
    "Mahdia", "Médenine", "Gafsa", "Jendouba",
]

STARTUP_NAMES = [
    "NovaSpark", "LumiPay", "TerraFarm", "MediPulse", "UrbanFlow",
    "ClearDoc", "SmartCraft", "EduBoost", "CarbonZero", "AgroMind",
    "HealthBridge", "TalentCore", "DataLens", "PayEase", "GreenVault",
    "CodeNest", "SolarLink", "MarketEye", "SkillUp", "BioTrack",
    "CloudFactory", "FoodRescue", "LearnPath", "EnergyPulse", "RoboFarm",
    "QuickCare", "HireAI", "CropSense", "FinGuru", "SafeRoute",
]


def _rand_bool(prob_true: float = 0.6) -> bool:
    return random.random() < prob_true


def _rand_range(options, weights=None) -> str:
    return random.choices(options, weights=weights, k=1)[0]


def generate_one(name: str = None, seed_override: int = None) -> dict:
    """Generate a single realistic startup application."""
    if seed_override is not None:
        random.seed(seed_override)

    age = _rand_range(
        AGE_OPTIONS,
        weights=[35, 35, 15, 5]   # Bias toward younger startups
    )

    # Legal creation probability varies by age
    legal_prob = {"Less than 2 years": 0.70, "2–5 years": 0.90, "5–7 years": 0.95, "More than 7 years": 0.98}
    legally_created = _rand_bool(legal_prob.get(age, 0.8))

    # Maturity correlates with age
    maturity_weights = {
        "Less than 2 years": [10, 20, 30, 25, 10, 4, 1],
        "2–5 years":         [2,  8,  20, 30, 25, 12, 3],
        "5–7 years":         [1,  3,  10, 20, 30, 25, 11],
        "More than 7 years": [0,  1,   5, 10, 25, 40, 19],
    }
    maturity = _rand_range(MATURITY_OPTIONS, weights=maturity_weights.get(age))

    # Maturity implies more clients / revenue
    mat_score = MATURITY_OPTIONS.index(maturity)
    rev_prob   = 0.05 + mat_score * 0.12
    client_w   = [max(1, 10 - mat_score * 2)] + [4, 4 + mat_score, 2 + mat_score, 1 + mat_score]
    emp_w      = [max(1, 8 - mat_score)] + [5, 3 + mat_score, 1 + mat_score // 2, 1]

    return {
        "startup_name":          name or random.choice(STARTUP_NAMES) + str(random.randint(1, 99)),
        "age":                   age,
        "sector":                random.choice(SECTORS),
        "governorate":           random.choice(GOVERNORATES),
        "legally_created":       legally_created,
        "full_time_founder":     _rand_bool(0.78),
        "founder_in_tunisia":    _rand_bool(0.82),
        "total_employees":       _rand_range(EMPLOYEE_RANGES, weights=emp_w),
        "salaried_employees":    _rand_range(EMPLOYEE_RANGES[:-1], weights=[4, 4, 3, 2]),
        "gender_mixed":          _rand_bool(0.52),
        "maturity":              maturity,
        "num_clients":           _rand_range(CLIENT_RANGES, weights=client_w),
        "generating_revenue":    _rand_bool(rev_prob),
        "startup_label":         _rand_bool(0.15),
        "raised_funding":        _rand_bool(0.18),
        "participated_programs": _rand_bool(0.22),
        "submitted_at":          (datetime.now() - timedelta(days=random.randint(0, 30))).strftime("%Y-%m-%d %H:%M"),
    }


def generate_batch(n: int = 30) -> list[dict]:
    """Generate n random applications."""
    random.seed(0)
    return [generate_one() for _ in range(n)]


# ── Google Form CSV format ───────────────────────────────────────────────────
GFORM_COLUMN_MAP = {
    "Timestamp":                        "submitted_at",
    "What is the name of your startup?": "startup_name",
    "Age of your startup":              "age",
    "Sector of your startup":           "sector",
    "Governorate":                      "governorate",
    "Is your startup legally established?":            "legally_created",
    "Is one of the founders working full-time?":       "full_time_founder",
    "Is at least one founder a resident of Tunisia?":  "founder_in_tunisia",
    "Total number of employees":         "total_employees",
    "Number of salaried employees":      "salaried_employees",
    "Is your team gender-mixed?":        "gender_mixed",
    "How advanced is your solution?":    "maturity",
    "Current number of clients/users":   "num_clients",
    "Is your startup currently generating revenue?": "generating_revenue",
    "Does your startup have a startup label?":       "startup_label",
    "Has your startup raised funding?":              "raised_funding",
    "Participated in acceleration programs?":        "participated_programs",
}

GFORM_BOOL_YES = {"Yes", "Oui", "TRUE", "True", "true", "1"}
GFORM_BOOL_NO  = {"No",  "Non", "FALSE","False","false","0"}


def apps_to_gform_csv(apps: list[dict]) -> str:
    """Serialise apps list to a Google-Form-style CSV string."""
    headers = list(GFORM_COLUMN_MAP.keys())
    inv_map = {v: k for k, v in GFORM_COLUMN_MAP.items()}

    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=headers)
    writer.writeheader()
    for app in apps:
        row = {}
        for gform_col, field in GFORM_COLUMN_MAP.items():
            val = app.get(field, "")
            if isinstance(val, bool):
                val = "Yes" if val else "No"
            row[gform_col] = val
        writer.writerow(row)
    return out.getvalue()


def gform_csv_to_apps(csv_text: str) -> list[dict]:
    """Parse a Google-Form CSV back into engine-ready dicts."""
    reader = csv.DictReader(io.StringIO(csv_text))
    apps = []
    for row in reader:
        app = {}
        for gform_col, field in GFORM_COLUMN_MAP.items():
            val = row.get(gform_col, "")
            # Coerce booleans
            if field in ("legally_created", "full_time_founder", "founder_in_tunisia",
                         "gender_mixed", "generating_revenue", "startup_label",
                         "raised_funding", "participated_programs"):
                app[field] = val in GFORM_BOOL_YES
            else:
                app[field] = val
        apps.append(app)
    return apps


# ── Typeform JSON format ──────────────────────────────────────────────────────
TYPEFORM_FIELD_IDS = {
    "startup_name":          "field_001",
    "age":                   "field_002",
    "sector":                "field_003",
    "governorate":           "field_004",
    "legally_created":       "field_005",
    "full_time_founder":     "field_006",
    "founder_in_tunisia":    "field_007",
    "total_employees":       "field_008",
    "salaried_employees":    "field_009",
    "gender_mixed":          "field_010",
    "maturity":              "field_011",
    "num_clients":           "field_012",
    "generating_revenue":    "field_013",
    "startup_label":         "field_014",
    "raised_funding":        "field_015",
    "participated_programs": "field_016",
}


def apps_to_typeform_json(apps: list[dict]) -> str:
    """Serialise apps list to a Typeform-style JSON string."""
    responses = []
    for i, app in enumerate(apps):
        answers = []
        for field, field_id in TYPEFORM_FIELD_IDS.items():
            val = app.get(field, "")
            if isinstance(val, bool):
                answers.append({"field": {"id": field_id, "ref": field}, "type": "boolean", "boolean": val})
            else:
                answers.append({"field": {"id": field_id, "ref": field}, "type": "text", "text": str(val)})
        responses.append({
            "response_id": f"resp_{i:04d}",
            "submitted_at": app.get("submitted_at", ""),
            "answers": answers,
        })
    return json.dumps({"total_items": len(responses), "items": responses}, indent=2)


def typeform_json_to_apps(json_text: str) -> list[dict]:
    """Parse Typeform JSON export into engine-ready dicts."""
    data = json.loads(json_text)
    bool_fields = {
        "legally_created", "full_time_founder", "founder_in_tunisia",
        "gender_mixed", "generating_revenue", "startup_label",
        "raised_funding", "participated_programs",
    }
    apps = []
    for item in data.get("items", []):
        app = {"submitted_at": item.get("submitted_at", "")}
        for ans in item.get("answers", []):
            ref = ans.get("field", {}).get("ref", "")
            if not ref:
                continue
            if ref in bool_fields:
                app[ref] = ans.get("boolean", False)
            else:
                app[ref] = ans.get("text", "")
        apps.append(app)
    return apps


if __name__ == "__main__":
    # Quick smoke test
    apps = generate_batch(5)
    csv_str = apps_to_gform_csv(apps)
    parsed  = gform_csv_to_apps(csv_str)
    print(f"Generated {len(apps)} apps → CSV → parsed {len(parsed)} apps ✓")

    tf_str  = apps_to_typeform_json(apps)
    parsed2 = typeform_json_to_apps(tf_str)
    print(f"Generated {len(apps)} apps → Typeform JSON → parsed {len(parsed2)} apps ✓")