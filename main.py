
import re
import time
import zipfile
import requests

from pathlib import Path
from urllib.parse import urljoin
from bs4 import BeautifulSoup

# --------------------------------
# Configuration
# --------------------------------

PAGE_URL = "https://download.companieshouse.gov.uk/en_output.html"

BASE_DIR = Path("companies_house_data")
ZIP_DIR = BASE_DIR / "zip_files"
CSV_DIR = BASE_DIR / "csv_files"

ZIP_DIR.mkdir(parents=True, exist_ok=True)
CSV_DIR.mkdir(parents=True, exist_ok=True)

session = requests.Session()

session.headers.update({
    "User-Agent": "CompanyDataResearch/1.0",
    "Accept": "text/html,application/zip,*/*"
})


# --------------------------------
# Find all multiple-file downloads
# --------------------------------

def find_download_links():

    print("Reading Companies House page...")

    response = session.get(PAGE_URL, timeout=60)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    pattern = re.compile(
        r"BasicCompanyData-\d{4}-\d{2}-\d{2}-part\d+_\d+\.zip",
        re.IGNORECASE
    )

    downloads = []

    for link in soup.find_all("a", href=True):

        href = link["href"]
        filename = href.split("/")[-1].split("?")[0]

        if pattern.fullmatch(filename):
            downloads.append(urljoin(PAGE_URL, href))

    downloads = sorted(set(downloads))

    if not downloads:
        raise RuntimeError(
            "No multiple-file company data ZIPs found."
        )

    print(f"Found {len(downloads)} ZIP files.")

    for url in downloads:
        print(" -", url.split("/")[-1])

    return downloads


# --------------------------------
# Download ZIP files
# --------------------------------

def download_file(url):

    filename = url.split("/")[-1]
    destination = ZIP_DIR / filename

    if destination.exists():
        print(f"Already downloaded: {filename}")
        return destination

    print(f"Downloading: {filename}")

    response = session.get(
        url,
        stream=True,
        timeout=(30, 180)
    )

    response.raise_for_status()

    with open(destination, "wb") as file:

        for chunk in response.iter_content(
            chunk_size=1024 * 1024
        ):
            if chunk:
                file.write(chunk)

    print(f"Saved: {destination}")

    return destination


# --------------------------------
# Safely extract CSV files
# --------------------------------

def extract_csv(zip_path):

    print(f"Extracting: {zip_path.name}")

    extracted = []

    with zipfile.ZipFile(zip_path, "r") as archive:

        for member in archive.infolist():

            if member.is_dir():
                continue

            # Only extract CSV files.
            if not member.filename.lower().endswith(".csv"):
                continue

            filename = Path(member.filename).name

            if not filename:
                continue

            destination = CSV_DIR / (
                zip_path.stem + "_" + filename
            )

            with archive.open(member) as source:
                with open(destination, "wb") as target:
                    while True:
                        chunk = source.read(1024 * 1024)
                        if not chunk:
                            break
                        target.write(chunk)

            extracted.append(destination)

            print("Extracted:", destination.name)

    return extracted


# --------------------------------
# Main process
# --------------------------------

def main():

    downloads = find_download_links()

    all_csv_files = []

    for index, url in enumerate(downloads, start=1):

        print(f"\nProcessing {index}/{len(downloads)}")

        try:
            zip_path = download_file(url)

            csv_files = extract_csv(zip_path)

            all_csv_files.extend(csv_files)

            time.sleep(0.5)

        except Exception as error:
            print("ERROR:", url, error)

    print("\n========== COMPLETED ==========")

    print("ZIP files found:", len(downloads))
    print("CSV files extracted:", len(all_csv_files))
    print("ZIP directory:", ZIP_DIR)
    print("CSV directory:", CSV_DIR)


if __name__ == "__main__":
    main()
