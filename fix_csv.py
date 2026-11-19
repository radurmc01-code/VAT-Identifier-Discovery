import csv
from pathlib import Path

RESULTS_FILE = Path(r"vat_candidates.csv")

HEADER_15 = [
    "company_name", "company_number", "ch_postcode", "website", "website_status",
    "vat_number", "validation", "registered_name", "registered_address",
    "hmrc_postcode", "name_similarity", "postcode_match", "match_verdict",
    "source_pages", "error",
]

with open(RESULTS_FILE, "r", encoding="utf-8-sig", newline="") as f:
    valid_rows = [row for row in csv.reader(f) if len(row) == 15]

with open(RESULTS_FILE, "w", encoding="utf-8-sig", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(HEADER_15)
    writer.writerows(valid_rows)

print(f"Reparat! Rânduri păstrate (15 coloane): {len(valid_rows)}")
