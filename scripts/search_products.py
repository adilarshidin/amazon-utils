import os
import csv
import json
import time
import requests
import numpy as np
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()

TOKEN = os.getenv("GOOGLE_CUSTOM_SEARCH_API_TOKEN")
ENGINE_ID = os.getenv("ENGINE_ID")
CSV_FILE = "input/konus.csv"
OUTPUT_CSV_FILE = "output/catalog_with_suggestions.csv"

COSTS_COLUMN = "TarifaCostos"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; PriceBot/1.0)"
}

FX_API = "https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies/eur.json"


# ---------- CSV ----------
def get_first_product(csv_file):
    with open(csv_file, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return next(reader)


# ---------- GOOGLE ----------
def google_search(query, start):
    url = "https://www.googleapis.com/customsearch/v1"
    params = {
        "key": TOKEN,
        "cx": ENGINE_ID,
        "q": query,
        "num": 10,
        "start": start,
    }
    r = requests.get(url, params=params)
    r.raise_for_status()
    return r.json().get("items", [])


# ---------- FX ----------
def load_fx_rates():
    r = requests.get(FX_API)
    r.raise_for_status()
    return r.json()["eur"]


def to_eur(price, currency, rates):
    currency = currency.lower()
    if currency == "eur":
        return price
    if currency not in rates:
        return None
    return price / rates[currency]


# ---------- Finance ----------

def compute_margin(price, cost):
    margin_eur = price - cost
    margin_pct = (margin_eur / cost) * 100 if cost > 0 else 0
    return margin_eur, margin_pct


def parse_money(value: str) -> float:
    """
    Converts values like:
    '4,80 €' -> 4.80
    '1.234,56 €' -> 1234.56
    '3 €' -> 3.0
    """
    if not value:
        return 0.0

    value = value.strip()

    # Remove currency symbols and spaces
    value = (
        value.replace("€", "")
             .replace("\xa0", "")
             .replace(" ", "")
    )

    # Handle thousands + decimal
    if "," in value and "." in value:
        # EU format: 1.234,56
        value = value.replace(".", "").replace(",", ".")
    else:
        # Simple decimal comma
        value = value.replace(",", ".")

    return float(value)


# ---------- TIER 1 ----------
def extract_from_pagemap(item):
    pagemap = item.get("pagemap", {})
    offers = pagemap.get("offer", [])

    if not offers:
        return None

    offer = offers[0]
    if "price" not in offer or "pricecurrency" not in offer:
        return None

    return float(offer["price"]), offer["pricecurrency"]


# ---------- TIER 2 ----------
def extract_from_page(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        r.raise_for_status()
    except Exception:
        return None

    soup = BeautifulSoup(r.text, "lxml")

    # JSON-LD
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string)
        except Exception:
            continue

        if isinstance(data, dict) and data.get("@type") == "Product":
            offers = data.get("offers", {})
            if isinstance(offers, dict):
                price = offers.get("price")
                currency = offers.get("priceCurrency")
                if price and currency:
                    return float(price), currency

    # Meta fallback
    mp = soup.find("meta", property="product:price:amount")
    mc = soup.find("meta", property="product:price:currency")

    if mp and mc:
        return float(mp["content"]), mc["content"]

    return None


# ---------- OUTLIERS ----------
def remove_outliers(prices):
    if len(prices) < 5:
        return prices

    prices = np.array(prices)
    low = np.percentile(prices, 10)
    high = np.percentile(prices, 90)

    return prices[(prices >= low) & (prices <= high)]


# ---------- MAIN WITH CHECKPOINT ----------
def main():
    # Read all rows from input CSV
    with open(CSV_FILE, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    # Load already processed EANs from output CSV
    processed_eans = set()
    if os.path.exists(OUTPUT_CSV_FILE):
        with open(OUTPUT_CSV_FILE, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                processed_eans.add(row["EAN"])

    fx_rates = load_fx_rates()

    # Prepare output CSV for appending
    output_exists = os.path.exists(OUTPUT_CSV_FILE)
    with open(OUTPUT_CSV_FILE, "a", newline="", encoding="utf-8") as f_out:
        fieldnames = list(rows[0].keys()) + ["Suggested Margin €", "Suggested Margin %", "Suggested PVP €"]
        writer = csv.DictWriter(f_out, fieldnames=fieldnames)

        if not output_exists:
            writer.writeheader()

        for product in rows:
            ean = product["EAN"]

            if ean in processed_eans:
                print(f"Skipping already processed product: {ean}")
                continue

            print(f"\nSearching Google for: {ean}\n")

            prices_eur = []

            for start in (1, 11, 21, 31):
                items = google_search(ean, start)

                for item in items:
                    data = extract_from_pagemap(item)

                    if data is None:
                        data = extract_from_page(item.get("link"))

                    if data:
                        price, currency = data
                        price_eur = to_eur(price, currency, fx_rates)
                        if price_eur:
                            prices_eur.append(price_eur)

                    time.sleep(0.4)

            if not prices_eur:
                print("No prices collected for this product.")
                # Still keep row, add empty suggestions
                product["Suggested Margin €"] = ""
                product["Suggested Margin %"] = ""
                product["Suggested PVP €"] = ""
                writer.writerow(product)
                processed_eans.add(ean)
                continue

            clean_prices = remove_outliers(prices_eur)

            # ---- READ SPREADSHEET VALUES ----
            cost = parse_money(product[COSTS_COLUMN])
            current_pvp = parse_money(product["Final PVP €"])

            # ---- CURRENT MARGIN ----
            current_margin_eur, current_margin_pct = compute_margin(
                current_pvp, cost
            )

            # ---- SUGGESTED MARGIN (USING MEDIAN MARKET PRICE + BUFFER) ----
            BUFFER = 0.10  # 10% above median
            market_price = float(np.median(clean_prices))
            suggested_pvp = market_price * (1 + BUFFER)
            suggested_margin_eur, suggested_margin_pct = compute_margin(
                suggested_pvp, cost
            )

            # ---- ADD SUGGESTION COLUMNS ----
            product["Suggested Margin €"] = f"{suggested_margin_eur:.2f}"
            product["Suggested Margin %"] = f"{suggested_margin_pct:.2f}"
            product["Suggested PVP €"] = f"{suggested_pvp:.2f}"

            writer.writerow(product)
            processed_eans.add(ean)

            print(f"Product {ean}: Suggested PVP € {suggested_pvp:.2f}, Margin € {suggested_margin_eur:.2f}, Margin % {suggested_margin_pct:.2f}")

    print(f"\n✅ Done! Results saved to {OUTPUT_CSV_FILE}")


if __name__ == "__main__":
    main()
