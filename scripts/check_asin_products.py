import pandas as pd
import requests
import time
import random
import os
from itertools import cycle
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


INPUT_FILE = "output/translated_catalog.csv"
OUTPUT_FILE = "output/translated_catalog_valid.csv"
CHECKPOINT_FILE = "input/asin_checkpoint.csv"

ASIN_COLUMN = "ASIN"
BASE_URL = "https://www.amazon.com/dp/{}"

# ====== OXYLABS PROXIES ======
PROXY_USERNAME = "USERNAME"  # fill if needed
PROXY_PASSWORD = "PASSWORD"  # fill if needed

PROXY_HOSTS = [
    "dc.oxylabs.io:8001",
    "dc.oxylabs.io:8002",
    "dc.oxylabs.io:8003",
    "dc.oxylabs.io:8004",
    "dc.oxylabs.io:8005",
]


def build_proxy(host):
    if PROXY_USERNAME and PROXY_PASSWORD:
        proxy = f"http://{PROXY_USERNAME}:{PROXY_PASSWORD}@{host}"
    else:
        proxy = f"http://{host}"

    return {"http": proxy, "https": proxy}

proxy_pool = cycle([build_proxy(h) for h in PROXY_HOSTS])

# ====== HEADERS ======
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/121.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Connection": "keep-alive",
}

# ====== SESSION ======
def create_session():
    session = requests.Session()
    session.headers.update(HEADERS)

    retries = Retry(
        total=4,
        backoff_factor=1.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"]
    )

    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    return session

# ====== AMAZON PAGE VALIDATION ======
def is_valid_amazon_product(response):
    if response.status_code == 404:
        return False

    text = response.text.lower()

    soft_404_signals = [
        "sorry! we couldn't find that page",
        "page not found",
        "looking for something?",
        "dogs of amazon",
        "error occurred",
        "enter the characters you see below"
    ]

    return not any(signal in text for signal in soft_404_signals)

# ====== MAIN ======
def main():
    df = pd.read_csv(INPUT_FILE)

    if ASIN_COLUMN not in df.columns:
        raise ValueError(f"Missing '{ASIN_COLUMN}' column")

    # Load checkpoint if exists
    if os.path.exists(CHECKPOINT_FILE):
        checkpoint_df = pd.read_csv(CHECKPOINT_FILE)
        processed_asins = set(checkpoint_df[ASIN_COLUMN])
        results = checkpoint_df.to_dict("records")
        print(f"🔁 Resuming from checkpoint ({len(processed_asins)} processed)")
    else:
        processed_asins = set()
        results = []

    session = create_session()

    for idx, row in df.iterrows():
        asin = str(row[ASIN_COLUMN]).strip()

        if asin in processed_asins:
            continue

        url = BASE_URL.format(asin)
        proxy = next(proxy_pool)

        try:
            response = session.get(
                url,
                proxies=proxy,
                timeout=12
            )

            valid = is_valid_amazon_product(response)

        except requests.RequestException:
            valid = False

        row_data = row.to_dict()
        row_data["__valid"] = valid
        results.append(row_data)

        # Write checkpoint immediately
        pd.DataFrame(results).to_csv(CHECKPOINT_FILE, index=False)

        status = "VALID" if valid else "REMOVED"
        print(f"[{len(results)}] {asin} → {status}")

        # Human-like delay
        time.sleep(random.uniform(2.0, 5.0))

    # Save final filtered file
    final_df = pd.DataFrame(results)
    final_df = final_df[final_df["__valid"] == True]
    final_df.drop(columns="__valid", inplace=True)

    final_df.to_csv(OUTPUT_FILE, index=False)

    print(f"\n✅ Done. Valid ASINs saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
