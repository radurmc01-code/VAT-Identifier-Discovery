import re
from difflib import SequenceMatcher
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR if BASE_DIR.name == "companies_house_data" else BASE_DIR / "companies_house_data"

INPUT_CSV = DATA_DIR / "results" / "vat_candidates.csv"
OUTPUT_CSV = DATA_DIR / "results" / "companies_by_vat.csv"

NAME_SIM_THRESHOLD = 0.80  # nume identic sau ușor diferit
VAT_SIM_THRESHOLD = 0.85   # VAT identic sau foarte asemănător


def clean_vat(vat: str) -> str:
    if not isinstance(vat, str):
        return ""
    return re.sub(r"[^A-Z0-9]", "", vat.upper()).strip()


def similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def main():
    df = pd.read_csv(INPUT_CSV, dtype=str).fillna("")

    # 1. Păstrăm DOAR rândurile care au un cod VAT extras
    df["clean_vat"] = df["vat_number"].apply(clean_vat)
    df = df[df["clean_vat"] != ""].copy()

    if df.empty:
        print("Nu există rânduri cu cod VAT extras în fișier.")
        return

    # Eliminăm duplicatele exacte pe (company_number, clean_vat)
    df = df.drop_duplicates(subset=["company_number", "clean_vat"]).reset_index(drop=True)

    # 2. Grupăm și organizăm firmele după codul VAT (păstrând numele original intact)
    vat_groups = (
        df.groupby("clean_vat", as_index=False)
        .agg(
            vat_number=("vat_number", "first"),
            companies_count=("company_name", "nunique"),
            company_names=("company_name", lambda s: " | ".join(dict.fromkeys(x for x in s if x))),
            company_numbers=("company_number", lambda s: " | ".join(dict.fromkeys(x for x in s if x))),
            validations=("validation", lambda s: " | ".join(dict.fromkeys(x for x in s if x))),
            verdicts=("match_verdict", lambda s: " | ".join(dict.fromkeys(x for x in s if x))),
        )
    )

    # 3. Verificăm și similaritatea între grupuri (nume identic/ușor diferit + VAT identic/asemănător)
    rows = vat_groups.to_dict(orient="records")
    notes = []

    for idx_a, g_a in enumerate(rows):
        group_notes = []
        for idx_b, g_b in enumerate(rows):
            if idx_a == idx_b:
                continue
            v_sim = similarity(g_a["clean_vat"], g_b["clean_vat"])
            n_sim = similarity(g_a["company_names"].upper(), g_b["company_names"].upper())
            if v_sim >= VAT_SIM_THRESHOLD and n_sim >= NAME_SIM_THRESHOLD:
                group_notes.append(
                    f"Similar VAT & name ({v_sim:.2f}/{n_sim:.2f}) with {g_b['vat_number']} [{g_b['company_names']}]"
                )
        if g_a["companies_count"] > 1:
            group_notes.append(f"Multiple companies ({g_a['companies_count']}) on same VAT")
        notes.append(" ; ".join(group_notes))

    vat_groups["flag_notes"] = notes
    vat_groups = vat_groups.sort_values(
        by=["companies_count", "clean_vat"], ascending=[False, True]
    ).drop(columns=["clean_vat"])

    vat_groups.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"Organizate {len(vat_groups)} coduri VAT (doar rânduri cu VAT extras). Salvat în: {OUTPUT_CSV}")
    print(vat_groups.head(20).to_string(index=False))


if __name__ == "__main__":
    main()