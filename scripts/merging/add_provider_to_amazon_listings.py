import pandas as pd
import os
import json
import re
from dotenv import load_dotenv
from mistralai import Mistral

# =========================
# CONFIG
# =========================
load_dotenv()

LISTINGS_CSV = "output/all_listings_with_images_and_category_translated.csv"
CATALOG_CSV = "input/catalog_initial.csv"
OUTPUT_CSV = "output/all_listings_ready.csv"
TMP_OUTPUT_FILE = "output/.all_listings_ready.tmp"

BATCH_SIZE = 15
LLM_MODEL = "mistral-small-latest"

# =========================
# HELPERS
# =========================
def clean_json(text: str):
    text = text.strip()
    text = re.sub(r"^```json\s*|\s*```$", "", text, flags=re.IGNORECASE)
    return json.loads(text)

def extract_manufacturer_batch(items, mistral):
    """
    Ask LLM to extract the provider (manufacturer) from product info
    items: list of dicts with keys: seller_sku, title, brand, description
    """
    prompt = f"""
You are extracting the PROVIDER of products for an e-commerce catalog.

Rules:
- For EACH product, return EXACTLY ONE provider name
- If unknown, guess based on brand, title, description
- Return ONLY a JSON array of strings, in the same order
- No markdown, no explanations

Products:
{json.dumps(items, ensure_ascii=False, indent=2)}
"""
    res = mistral.chat.complete(
        model=LLM_MODEL,
        messages=[{"role": "user", "content": prompt}],
        stream=False,
    )
    return clean_json(res.choices[0].message.content)

# =========================
# LOAD CSV FILES
# =========================
listings_df = pd.read_csv(LISTINGS_CSV, dtype=str)
catalog_df = pd.read_csv(CATALOG_CSV, dtype=str)

# Keep only the latest entry per EAN
catalog_df['FECHA'] = pd.to_datetime(catalog_df['FECHA'], errors='coerce')
catalog_df = catalog_df.sort_values('FECHA').drop_duplicates(subset='EAN', keep='last')

# Ensure merge keys are strings
listings_df['seller-sku'] = listings_df['seller-sku'].astype(str)
catalog_df['EAN'] = catalog_df['EAN'].astype(str)

# Merge PROVEEDOR from catalog
catalog_subset = catalog_df[['EAN', 'PROVEEDOR']]
merged_df = listings_df.merge(
    catalog_subset,
    how='left',
    left_on='seller-sku',
    right_on='EAN'
)
merged_df = merged_df.drop(columns=['EAN'])

# =========================
# LLM FALLBACK FOR MISSING PROVEEDOR
# =========================
unmatched_indexes = merged_df[
    merged_df['PROVEEDOR'].isna() | (merged_df['PROVEEDOR'] == "")
].index.tolist()

print(f"🤖 LLM needed for {len(unmatched_indexes)} rows")

with Mistral(api_key=os.getenv("MISTRAL_API_TOKEN", "")) as mistral:
    for i in range(0, len(unmatched_indexes), BATCH_SIZE):
        batch_indexes = unmatched_indexes[i:i + BATCH_SIZE]

        items = []
        for idx in batch_indexes:
            row = merged_df.loc[idx]
            items.append({
                "seller_sku": row.get("seller-sku", ""),
                "title": row.get("item-name", ""),
                "brand": row.get("brand-name", ""),
                "description": row.get("item-description", "")
            })

        guessed_manufacturers = extract_manufacturer_batch(items, mistral)

        for df_idx, manufacturer in zip(batch_indexes, guessed_manufacturers):
            merged_df.at[df_idx, "PROVEEDOR"] = manufacturer

        # Atomic write per batch
        merged_df.to_csv(TMP_OUTPUT_FILE, index=False)
        os.replace(TMP_OUTPUT_FILE, OUTPUT_CSV)

        print(f"✅ LLM processed rows {i + 1}–{min(i + BATCH_SIZE, len(unmatched_indexes))}")

# =========================
# FINALIZE OUTPUT
# =========================
merged_df = merged_df.rename(columns={"PROVEEDOR": "manufacturer"})

# Remove duplicates if any
merged_df = merged_df.drop_duplicates(subset=['seller-sku', 'manufacturer'])

# Write final output
merged_df.to_csv(OUTPUT_CSV, index=False)

print(f"🎉 Completed. Final output: {OUTPUT_CSV}")
