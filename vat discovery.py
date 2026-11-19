import os
import random
import re
import time
import zipfile
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from dotenv import find_dotenv, load_dotenv, set_key

# =====================================================
# CONFIGURATION
# =====================================================

BASE_DIR = Path(__file__).resolve().parent

ENV_PATH = BASE_DIR / ".env"
if not ENV_PATH.exists():
    ENV_PATH = BASE_DIR.parent / ".env"

if ENV_PATH.exists():
    load_dotenv(ENV_PATH, override=False)
else:
    load_dotenv(find_dotenv(), override=False)

BRAVE_API_KEY = os.getenv("BRAVE_API_KEY", "").strip()
if not BRAVE_API_KEY:
    print(f"BRAVE_API_KEY nu a fost găsită în mediu sau în {ENV_PATH}.")
    BRAVE_API_KEY = input("Introdu BRAVE_API_KEY aici și apasă Enter: ").strip()
    if not BRAVE_API_KEY:
        raise ValueError("BRAVE_API_KEY este obligatorie pentru a rula scriptul.")

    ENV_PATH.touch(exist_ok=True)
    set_key(str(ENV_PATH), "BRAVE_API_KEY", BRAVE_API_KEY)
    print(f"Cheia a fost salvată automat în {ENV_PATH}.")

# Detectează automat dacă scriptul e în companies_house_data sau în rădăcină
DATA_DIR = BASE_DIR if BASE_DIR.name == "companies_house_data" else BASE_DIR / "companies_house_data"
CSV_INPUT_DIR = DATA_DIR / "csv_files"
ZIP_INPUT_DIR = DATA_DIR / "zip_files"

OUTPUT_FOLDER = DATA_DIR / "results"
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

SAMPLE_FILE = OUTPUT_FOLDER / "sample_companies.csv"
PROGRESS_FILE = OUTPUT_FOLDER / "progress.csv"
RESULTS_FILE = OUTPUT_FOLDER / "vat_candidates.csv"
REPORT_FILE = OUTPUT_FOLDER / "final_report.xlsx"

CHUNK_SIZE = 20000
SAMPLE_SIZE = 200
RANDOM_SEED = 150
ONLY_ACTIVE = True

MAX_PAGES = 6
REQUEST_DELAY = 0.5
BRAVE_DELAY = 1.1
HMRC_DELAY = 0.5

VIES_URL = "https://ec.europa.eu/taxation_customs/vies/rest-api/check-vat-number"
HMRC_URL = "https://api.service.hmrc.gov.uk/organisations/vat/check-vat-number/lookup"
SEARCH_URL = "https://api.search.brave.com/res/v1/web/search"

RESULT_COLUMNS = [
    "company_name",
    "company_number",
    "ch_postcode",
    "website",
    "website_status",
    "vat_number",
    "validation",
    "registered_name",
    "registered_address",
    "hmrc_postcode",
    "name_similarity",
    "postcode_match",
    "match_verdict",
    "source_pages",
    "error",
]

session = requests.Session()
session.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/130.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-GB,en;q=0.9",
})

# =====================================================
# VAT PATTERNS & CONSTANTS
# =====================================================

VAT_PATTERNS = [
    r"\bGB\s?\d{3}\s?\d{4}\s?\d{2}(?:\s?\d{3})?\b",
    r"\bGB\s?(?:GD|HA)\s?\d{3}\b",
    r"\bATU\d{8}\b",
    r"\bBE0?\d{9}\b",
    r"\bBG\d{9,10}\b",
    r"\bCY\d{8}[A-Z]\b",
    r"\bCZ\d{8,10}\b",
    r"\bDE\d{9}\b",
    r"\bDK\d{8}\b",
    r"\bEE\d{9}\b",
    r"\bEL\d{9}\b",
    r"\bES[A-Z0-9]\d{7}[A-Z0-9]\b",
    r"\bFI\d{8}\b",
    r"\bFR[A-Z0-9]{2}\d{9}\b",
    r"\bHR\d{11}\b",
    r"\bHU\d{8}\b",
    r"\bIE\d{7}[A-Z]{1,2}\b",
    r"\bIT\d{11}\b",
    r"\bLT\d{9,12}\b",
    r"\bLU\d{8}\b",
    r"\bLV\d{11}\b",
    r"\bMT\d{8}\b",
    r"\bNL\d{9}B\d{2}\b",
    r"\bPL\d{10}\b",
    r"\bPT\d{9}\b",
    r"\bRO\d{6,10}\b",
    r"\bSE\d{12}\b",
    r"\bSI\d{8}\b",
    r"\bSK\d{10}\b",
    r"\bXI\d{9}(?:\d{3})?\b",
]

COMPILED_VAT_PATTERNS = [re.compile(p, re.IGNORECASE) for p in VAT_PATTERNS]

# Pattern contextual pentru coduri UK VAT scrise FĂRĂ prefixul "GB"
# ex: "VAT No: 123 4567 89", "VAT Registration Number 123456789", "VAT Reg: 123-4567-89"
UK_CONTEXTUAL_VAT_PATTERN = re.compile(
    r"\bVAT\s*(?:Reg(?:istration)?\.?\s*)?(?:No\.?|Number|ID|Ref\.?)?\s*[:#\-]?\s*"
    r"(?:GB\s*)?(\d{3}[\s\-]?\d{4}[\s\-]?\d{2}(?:[\s\-]?\d{3})?)\b",
    re.IGNORECASE,
)

EU_CODES = {
    "AT", "BE", "BG", "CY", "CZ", "DE", "DK",
    "EE", "EL", "ES", "FI", "FR", "HR", "HU",
    "IE", "IT", "LT", "LU", "LV", "MT", "NL",
    "PL", "PT", "RO", "SE", "SI", "SK", "XI",
}

EXCLUDED_DOMAINS = {
    "companieshouse.gov.uk",
    "gov.uk",
    "nhs.uk",
    "thegazette.co.uk",
    "facebook.com",
    "linkedin.com",
    "instagram.com",
    "twitter.com",
    "x.com",
    "youtube.com",
    "wikipedia.org",
    "opencorporates.com",
    "endole.co.uk",
    "pomanda.com",
    "companycheck.co.uk",
    "duedil.com",
    "creditsafe.com",
    "bizdb.co.uk",
    "tracxn.com",
    "findmesales.com",
    "greatbusinessexchange.co.uk",
    "gbrbusiness.com",
    "yell.com",
    "cylex-uk.co.uk",
    "192.com",
    "bloomberg.com",
    "dnb.com",
    "zoominfo.com",
    "crunchbase.com",
}

PAGE_KEYWORDS = (
    "contact", "about", "legal", "terms",
    "privacy", "imprint", "company", "vat", "footer",
)


# =====================================================
# HELPERS
# =====================================================

def normalize_column(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def clean_value(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_vat(value: str) -> str:
    return re.sub(r"[\s.\-]", "", value.upper())


def normalize_postcode(postcode: str) -> str:
    return re.sub(r"\s+", "", postcode.upper()).strip()


def normalize_company_name(name: str) -> str:
    text = name.upper()
    text = re.sub(r"&", " AND ", text)
    text = re.sub(r"\b(LIMITED|LTD|PLC|LLP|INC|UK|GROUP|HOLDINGS|SERVICES|SVCS)\b\.?", "", text)
    text = re.sub(r"[^A-Z0-9\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def compute_name_similarity(name_a: str, name_b: str) -> float:
    norm_a = normalize_company_name(name_a)
    norm_b = normalize_company_name(name_b)
    if not norm_a or not norm_b:
        return 0.0
    return round(SequenceMatcher(None, norm_a, norm_b).ratio(), 3)


def evaluate_match(
    validation: str,
    ch_name: str,
    ch_postcode: str,
    registered_name: str,
    hmrc_postcode: str,
) -> tuple[float, bool, str]:
    """
    Compară firma din Companies House cu entitatea returnată de HMRC/VIES.
    Returnează (name_similarity, postcode_match, match_verdict).
    """
    if validation not in ("UK_VALID_HMRC", "VALID"):
        return 0.0, False, "N/A"

    if not registered_name:
        return 0.0, False, "NEEDS_MANUAL_REVIEW"

    similarity = compute_name_similarity(ch_name, registered_name)
    pc_match = (
        bool(ch_postcode)
        and bool(hmrc_postcode)
        and normalize_postcode(ch_postcode) == normalize_postcode(hmrc_postcode)
    )

    if similarity >= 0.75 or (pc_match and similarity >= 0.50):
        return similarity, pc_match, "TRUE_POSITIVE"

    return similarity, pc_match, "FALSE_POSITIVE_MISMATCH"


def get_existing_companies() -> set[str]:
    if not PROGRESS_FILE.exists():
        return set()
    try:
        df = pd.read_csv(PROGRESS_FILE, dtype=str)
        assert isinstance(df, pd.DataFrame)
        return set(df["company_number"].dropna().tolist())
    except (OSError, KeyError, pd.errors.EmptyDataError, pd.errors.ParserError):
        return set()


def save_result(data: dict) -> None:
    file_exists = RESULTS_FILE.exists()
    row = {col: data.get(col, "") for col in RESULT_COLUMNS}
    pd.DataFrame([row], columns=RESULT_COLUMNS).to_csv(
        RESULTS_FILE,
        mode="a",
        header=not file_exists,
        index=False,
        encoding="utf-8-sig",
    )


def save_progress(company_number: str, company_name: str) -> None:
    file_exists = PROGRESS_FILE.exists()
    pd.DataFrame([{
        "company_number": company_number,
        "company_name": company_name,
        "processed": True,
    }]).to_csv(
        PROGRESS_FILE,
        mode="a",
        header=not file_exists,
        index=False,
        encoding="utf-8-sig",
    )


def request_with_retry(
    method: str,
    url: str,
    retries: int = 3,
    raise_for_status: bool = True,
    **kwargs,
) -> requests.Response:
    response: requests.Response | None = None
    for attempt in range(max(1, retries)):
        response = session.request(method, url, **kwargs)
        if response.status_code == 429:
            wait = 2 ** (attempt + 1)
            print(f"Rate limit (429) on {url}. Waiting {wait}s...")
            time.sleep(wait)
            continue
        if raise_for_status:
            response.raise_for_status()
        return response

    assert response is not None
    if raise_for_status:
        response.raise_for_status()
    return response


# =====================================================
# UNIFORM RANDOM SAMPLING (RESERVOIR SAMPLING)
# =====================================================

def iter_companies_house_chunks():
    csv_files = sorted(CSV_INPUT_DIR.glob("*.csv")) if CSV_INPUT_DIR.exists() else []
    zip_files = sorted(ZIP_INPUT_DIR.glob("*.zip")) if (not csv_files and ZIP_INPUT_DIR.exists()) else []

    if not csv_files and not zip_files:
        raise FileNotFoundError(
            f"Nu am găsit fișiere CSV în {CSV_INPUT_DIR} sau ZIP în {ZIP_INPUT_DIR}. Rulează main.py mai întâi."
        )

    for path in csv_files:
        yield path.name, pd.read_csv(
            path, dtype=str, chunksize=CHUNK_SIZE, encoding="utf-8-sig", on_bad_lines="skip"
        )

    for path in zip_files:
        with zipfile.ZipFile(path, "r") as archive:
            for member in archive.namelist():
                if not member.lower().endswith(".csv"):
                    continue
                with archive.open(member) as file:
                    yield member, pd.read_csv(
                        file, dtype=str, chunksize=CHUNK_SIZE, encoding="utf-8-sig", on_bad_lines="skip"
                    )


def load_or_create_sample() -> list[dict[str, str]]:
    """
    Încarcă eșantionul existent din sample_companies.csv sau generează unul nou
    uniform aleatoriu (Reservoir Sampling) din toate firmele active.
    """
    if SAMPLE_FILE.exists():
        df = pd.read_csv(SAMPLE_FILE, dtype=str)
        assert isinstance(df, pd.DataFrame)
        df = df.fillna("")
        print(f"Loaded existing sample ({len(df)} companies) from {SAMPLE_FILE.name}")
        return df.to_dict(orient="records")

    print(f"Building uniform random sample of {SAMPLE_SIZE} active companies (seed={RANDOM_SEED})...")
    rng = random.Random(RANDOM_SEED)
    reservoir: list[dict[str, str]] = []
    seen_active = 0

    for filename, reader in iter_companies_house_chunks():
        print(f"  Scanning {filename}...")
        for chunk in reader:
            columns = {normalize_column(col): col for col in chunk.columns}
            name_col = columns.get("companyname")
            number_col = columns.get("companynumber")
            status_col = columns.get("companystatus")
            postcode_col = columns.get("regaddresspostcode")

            if not name_col or not number_col:
                break

            if ONLY_ACTIVE and status_col:
                chunk = chunk[chunk[status_col].fillna("").str.strip().str.lower() == "active"]

            for _, row in chunk.iterrows():
                name = clean_value(row[name_col])
                number = clean_value(row[number_col])
                if not name or not number:
                    continue

                postcode = clean_value(row[postcode_col]) if postcode_col else ""
                seen_active += 1

                record = {
                    "company_name": name,
                    "company_number": number,
                    "ch_postcode": postcode,
                }

                if len(reservoir) < SAMPLE_SIZE:
                    reservoir.append(record)
                else:
                    idx = rng.randint(0, seen_active - 1)
                    if idx < SAMPLE_SIZE:
                        reservoir[idx] = record

    pd.DataFrame(reservoir).to_csv(SAMPLE_FILE, index=False, encoding="utf-8-sig")
    print(f"Sample saved to {SAMPLE_FILE} (drawn from {seen_active:,} active companies).")
    return reservoir


# =====================================================
# WEBSITE DISCOVERY
# =====================================================

def search_website(company_name: str) -> list[dict]:
    clean_name = normalize_company_name(company_name) or company_name
    query = f'"{clean_name}" UK'

    try:
        response = request_with_retry(
            "GET",
            SEARCH_URL,
            headers={"X-Subscription-Token": BRAVE_API_KEY},
            params={
                "q": query,
                "country": "GB",
                "search_lang": "en",
                "count": 10,
            },
            timeout=20,
        )
        time.sleep(BRAVE_DELAY)

        results = response.json().get("web", {}).get("results", [])
        candidates = []
        seen_domains: set[str] = set()

        for item in results:
            url = str(item.get("url", ""))
            parsed = urlparse(url)
            domain = parsed.netloc.lower()

            if not domain or domain in seen_domains:
                continue

            if parsed.scheme not in ("https", "http"):
                continue

            if any(
                domain == ex or domain.endswith("." + ex)
                for ex in EXCLUDED_DOMAINS
            ):
                continue

            seen_domains.add(domain)
            candidates.append({
                "website": url,
                "domain": domain,
                "title": item.get("title", ""),
                "description": item.get("description", ""),
            })

        return candidates

    except requests.RequestException as error:
        print(f"Search error: {error}")
        return []


# =====================================================
# WEBSITE SCANNING
# =====================================================

def extract_visible_text(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()
    return soup.get_text(" ", strip=True)


def extract_vat_numbers(text: str) -> set[str]:
    found: set[str] = set()

    # 1. Coduri explicite cu prefix de țară (GB..., DE..., FR..., etc.)
    for pattern in COMPILED_VAT_PATTERNS:
        for match in pattern.findall(text):
            found.add(clean_vat(match))

    # 2. Coduri UK VAT fără prefix "GB", dar precedate de cuvântul "VAT"
    for digits_match in UK_CONTEXTUAL_VAT_PATTERN.findall(text):
        digits = clean_vat(digits_match)
        if digits.isdigit() and len(digits) in (9, 12):
            found.add(f"GB{digits}")

    return found


def scan_website(website: str, company_number: str) -> tuple[list[str], list[str], bool, str]:
    found_vats: set[str] = set()
    scanned_pages: list[str] = []
    company_confirmed = False

    try:
        response = session.get(website, timeout=15)
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "")
        if "html" not in content_type.lower():
            return [], [], False, "NON_HTML_HOMEPAGE"

        base_url = str(response.url)
        base_domain = urlparse(base_url).netloc

        soup = BeautifulSoup(response.text, "html.parser")

        subpages: list[str] = []
        for link in soup.find_all("a", href=True):
            raw_href = link.get("href")
            if not isinstance(raw_href, str):
                continue

            target = urljoin(base_url, raw_href).split("#")[0]
            if urlparse(target).netloc != base_domain or target == base_url:
                continue

            link_text = (link.get_text(" ", strip=True) + " " + target).lower()
            if any(word in link_text for word in PAGE_KEYWORDS):
                if target not in subpages:
                    subpages.append(target)

            if len(subpages) >= MAX_PAGES - 1:
                break

        home_text = extract_visible_text(response.text)
        found_vats.update(extract_vat_numbers(home_text))
        if company_number and company_number in home_text:
            company_confirmed = True
        scanned_pages.append(base_url)

    except requests.RequestException as error:
        return [], [], False, str(error)

    for page in subpages:
        try:
            time.sleep(REQUEST_DELAY)
            resp = session.get(page, timeout=15)
            resp.raise_for_status()

            if "html" not in resp.headers.get("Content-Type", "").lower():
                continue

            text = extract_visible_text(resp.text)
            found_vats.update(extract_vat_numbers(text))
            if company_number and company_number in text:
                company_confirmed = True
            scanned_pages.append(page)

        except requests.RequestException:
            continue

    return sorted(found_vats), scanned_pages, company_confirmed, ""


# =====================================================
# VAT VALIDATION (HMRC LIVE + UK MOD-97 + EU VIES)
# =====================================================

def is_valid_uk_vat_checksum(digits: str) -> bool:
    if len(digits) not in (9, 12) or not digits.isdigit():
        return False

    base = digits[:9]
    weights = [8, 7, 6, 5, 4, 3, 2]
    total = sum(int(base[i]) * weights[i] for i in range(7))
    check = int(base[7:9])

    rem_old = (total + check) % 97
    rem_new = (total + check + 55) % 97
    return rem_old == 0 or rem_new == 0


def format_hmrc_address(address_obj: dict) -> tuple[str, str]:
    if not isinstance(address_obj, dict):
        return "", ""
    parts = [
        str(address_obj.get(k, "")).strip()
        for k in ("line1", "line2", "line3", "line4", "line5", "postcode", "countryCode")
        if address_obj.get(k)
    ]
    postcode = str(address_obj.get("postcode", "")).strip()
    return ", ".join(parts), postcode


def validate_hmrc(vrn: str) -> tuple[str, str, str, str]:
    """Returnează (status, registered_name, registered_address, hmrc_postcode)."""
    if vrn.startswith(("GD", "HA")):
        return "UK_GOV_OR_HEALTH_NOT_IN_HMRC_API", "", "", ""

    if not is_valid_uk_vat_checksum(vrn):
        return "UK_CHECKSUM_INVALID", "", "", ""

    url = f"{HMRC_URL}/{vrn}"

    try:
        response = request_with_retry(
            "GET",
            url,
            raise_for_status=False,
            headers={"Accept": "application/json"},
            timeout=20,
        )
        time.sleep(HMRC_DELAY)

        if response.status_code == 200:
            target = response.json().get("target", {})
            registered_name = str(target.get("name", "")).strip()
            registered_address, postcode = format_hmrc_address(target.get("address", {}))
            return "UK_VALID_HMRC", registered_name, registered_address, postcode

        if response.status_code == 404:
            return "UK_INVALID_HMRC", "", "", ""

        if response.status_code == 400:
            return "UK_INVALID_FORMAT_HMRC", "", "", ""

        return f"HMRC_HTTP_{response.status_code}", "", "", ""

    except requests.RequestException as error:
        print(f"HMRC error for GB{vrn}: {error}")
        return "HMRC_ERROR", "", "", ""


def validate_vat(vat: str) -> tuple[str, str, str, str]:
    """Returnează (validation_status, registered_name, registered_address, postcode)."""
    vat = clean_vat(vat)
    match = re.match(r"^([A-Z]{2})([A-Z0-9]+)$", vat)
    if not match:
        return "UNRECOGNIZED_FORMAT", "", "", ""

    country, number = match.group(1), match.group(2)

    if country == "GB":
        return validate_hmrc(number)

    if country == "GR":
        country = "EL"

    if country not in EU_CODES:
        return "NOT_ELIGIBLE_FOR_VIES", "", "", ""

    try:
        response = request_with_retry(
            "POST",
            VIES_URL,
            json={"countryCode": country, "vatNumber": number},
            timeout=30,
        )
        data = response.json()
        if "valid" in data:
            status = "VALID" if data["valid"] else "INVALID"
            registered_name = str(data.get("name", "")).strip() if data["valid"] else ""
            registered_address = str(data.get("address", "")).strip() if data["valid"] else ""
            if registered_name == "---":
                registered_name = ""
            if registered_address == "---":
                registered_address = ""
            return status, registered_name, registered_address, ""
        return "VIES_ERROR", "", "", ""

    except requests.RequestException as error:
        print(f"VIES error for {vat}: {error}")
        return "VIES_ERROR", "", "", ""


# =====================================================
# COMPANY PROCESSING
# =====================================================

def process_company(company_name: str, company_number: str, ch_postcode: str) -> None:
    print(f"\nCompany: {company_name} ({company_number}) [{ch_postcode}]")

    websites = search_website(company_name)
    if not websites:
        save_result({
            "company_name": company_name,
            "company_number": company_number,
            "ch_postcode": ch_postcode,
            "website_status": "NOT_FOUND",
            "validation": "NO_WEBSITE_FOUND",
            "match_verdict": "N/A",
        })
        return

    for candidate in websites[:3]:
        website = candidate["website"]
        print(f"  Candidate: {website}")

        vats, pages, confirmed, error = scan_website(website, company_number)
        site_status = "CONFIRMED" if confirmed else "CANDIDATE_NOT_CONFIRMED"

        if not vats:
            save_result({
                "company_name": company_name,
                "company_number": company_number,
                "ch_postcode": ch_postcode,
                "website": website,
                "website_status": site_status,
                "validation": "SCAN_ERROR" if error else "NO_VAT_FOUND",
                "match_verdict": "N/A",
                "source_pages": "; ".join(pages),
                "error": error,
            })
            continue

        found_true_positive = False
        for vat in vats:
            status, reg_name, reg_addr, hmrc_pc = validate_vat(vat)
            similarity, pc_match, verdict = evaluate_match(
                status, company_name, ch_postcode, reg_name, hmrc_pc
            )
            if verdict == "TRUE_POSITIVE":
                found_true_positive = True

            save_result({
                "company_name": company_name,
                "company_number": company_number,
                "ch_postcode": ch_postcode,
                "website": website,
                "website_status": site_status,
                "vat_number": vat,
                "validation": status,
                "registered_name": reg_name,
                "registered_address": reg_addr,
                "hmrc_postcode": hmrc_pc,
                "name_similarity": similarity,
                "postcode_match": pc_match,
                "match_verdict": verdict,
                "source_pages": "; ".join(pages),
            })
            print(
                f"    VAT: {vat} => {status} | {verdict} "
                f"(sim={similarity}, pc_match={pc_match}) | HMRC: {reg_name}"
            )

        if found_true_positive:
            break


# =====================================================
# REPORT
# =====================================================

def generate_report() -> None:
    if not RESULTS_FILE.exists():
        print("No results generated.")
        return

    df = pd.read_csv(RESULTS_FILE, dtype=str)
    assert isinstance(df, pd.DataFrame)
    df = df.fillna("")

    counts = df["validation"].value_counts()
    verdicts = df["match_verdict"].value_counts()

    tp = int(verdicts.get("TRUE_POSITIVE", 0))
    fp = int(verdicts.get("FALSE_POSITIVE_MISMATCH", 0))
    total_checked_matches = tp + fp
    fp_rate = f"{(fp / total_checked_matches) * 100:.1f}%" if total_checked_matches > 0 else "N/A"

    unique_companies = df["company_number"].nunique()
    companies_with_tp = df[df["match_verdict"] == "TRUE_POSITIVE"]["company_number"].nunique()
    coverage = f"{(companies_with_tp / unique_companies) * 100:.1f}%" if unique_companies > 0 else "0.0%"

    summary = pd.DataFrame({
        "Metric": [
            "Unique Companies Processed",
            "Companies with Confirmed True Positive VAT",
            "Sample Coverage (%)",
            "True Positives (Valid VAT + Entity Match)",
            "False Positives (Valid VAT + Wrong Entity)",
            "False Positive Rate (%)",
            "UK Valid VAT (HMRC)",
            "UK Invalid VAT (HMRC)",
            "UK Checksum Invalid (Local)",
            "EU Valid VAT (VIES)",
            "No VAT Found (Candidates)",
            "No Website Found",
            "Scan Errors",
        ],
        "Value": [
            unique_companies,
            companies_with_tp,
            coverage,
            tp,
            fp,
            fp_rate,
            int(counts.get("UK_VALID_HMRC", 0)),
            int(counts.get("UK_INVALID_HMRC", 0)),
            int(counts.get("UK_CHECKSUM_INVALID", 0)),
            int(counts.get("VALID", 0)),
            int(counts.get("NO_VAT_FOUND", 0)),
            int(counts.get("NO_WEBSITE_FOUND", 0)),
            int(counts.get("SCAN_ERROR", 0)),
        ],
    })

    with pd.ExcelWriter(REPORT_FILE, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        df.to_excel(writer, sheet_name="All Results", index=False)

    print("\n========== FINAL REPORT ==========")
    print(summary.to_string(index=False))


# =====================================================
# MAIN
# =====================================================

def main() -> None:
    sample = load_or_create_sample()
    processed = get_existing_companies()

    print("\nVAT IDENTIFIER DISCOVERY (HMRC + VIES)")
    print("======================================")
    print(f"Sample size: {len(sample)} | Already processed: {len(processed)}")

    run_count = 0
    for idx, company in enumerate(sample, start=1):
        number = company["company_number"]
        name = company["company_name"]
        postcode = company.get("ch_postcode", "")

        if number in processed:
            continue

        print(f"\n[{idx}/{len(sample)}]", end="")
        try:
            process_company(name, number, postcode)
            save_progress(number, name)
            processed.add(number)
            run_count += 1
        except (requests.RequestException, OSError, ValueError) as error:
            print(f"Processing error {number}: {error}")

    print(f"\nProcessing finished. Companies processed this run: {run_count}")
    generate_report()


if __name__ == "__main__":
    main()