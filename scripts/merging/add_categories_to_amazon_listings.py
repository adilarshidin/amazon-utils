import pandas as pd
import glob
import os
import json
import re
import signal
from dotenv import load_dotenv
from mistralai import Mistral

# =========================
# CONFIG
# =========================
load_dotenv()

INPUT_CSV = "output/all_listings_with_images.csv"
OUTPUT_CSV = "output/all_listings_with_images_and_category.csv"
TMP_OUTPUT_FILE = "output/.all_listings_with_images_and_category.tmp"
CHECKPOINT_FILE = "checkpoints/category_guess_checkpoint.txt"

XLSM_DIR = "input"

BATCH_SIZE = 15
LLM_MODEL = "mistral-small-latest"

# =========================
# HELPERS
# =========================
def clean_json(text: str):
    text = text.strip()
    text = re.sub(r"^```json\s*|\s*```$", "", text, flags=re.IGNORECASE)
    return json.loads(text)

def guess_categories_batch(items, allowed_categories, mistral):
    prompt = f"""
You are classifying Amazon catalog products.

Rules:
- For EACH item, select EXACTLY ONE category
- Category MUST be one of the allowed categories list
- Use product context to choose best fit
- Return ONLY valid JSON array
- Keep same order
- No markdown, no explanations

Allowed categories:
{allowed_categories}

Products to classify:
{json.dumps(items, ensure_ascii=False, indent=2)}

Return format example:
["Category A", "Category B", ...]
"""
    res = mistral.chat.complete(
        model=LLM_MODEL,
        messages=[{"role": "user", "content": prompt}],
        stream=False,
    )

    return clean_json(res.choices[0].message.content)

# Graceful Ctrl+C
stop_requested = False
def handle_sigint(signum, frame):
    global stop_requested
    stop_requested = True
    print("\n🛑 Ctrl+C detected. Finishing current batch safely...")

signal.signal(signal.SIGINT, handle_sigint)

# =========================
# LOAD MAIN CSV
# =========================
df = pd.read_csv(INPUT_CSV, dtype=str)
df["amazon_tipo_de_producto"] = df.get("amazon_tipo_de_producto", "")

# =========================
# BUILD SKU → CATEGORY MAP
# + COLLECT ALL CATEGORIES
# =========================
sku_to_category = {}
all_categories = set()

xlsm_files = glob.glob(os.path.join(XLSM_DIR, "*.xlsm"))
print(f"📂 Found {len(xlsm_files)} XLSM files")

for file in xlsm_files:
    print(f"🔍 Reading {file}")
    try:
        sheet = pd.read_excel(
            file,
            sheet_name="Plantilla",
            header=3,
            dtype=str
        )
        sheet = sheet.iloc[3:]

        if not {"SKU", "Tipo de producto"}.issubset(sheet.columns):
            print(f"⚠️ Missing columns in {file}")
            continue

        for _, row in sheet.iterrows():
            sku = row.get("SKU")
            tipo = row.get("Tipo de producto")

            if pd.notna(tipo):
                all_categories.add(tipo)

            if pd.notna(sku) and pd.notna(tipo):
                if sku not in sku_to_category:
                    sku_to_category[sku] = tipo

    except Exception as e:
        print(f"❌ Error reading {file}: {e}")

all_categories = sorted(all_categories)
print(f"🏷️ Loaded {len(all_categories)} distinct Amazon categories")

# =========================
# APPLY DETERMINISTIC MATCH
# =========================
matched = 0
for idx, row in df.iterrows():
    sku = row.get("seller-sku")
    if sku in sku_to_category:
        df.at[idx, "amazon_tipo_de_producto"] = sku_to_category[sku]
        matched += 1

print(f"🎯 Deterministic matched {matched}/{len(df)}")

# =========================
# LLM FALLBACK FOR UNMATCHED
# =========================
unmatched_indexes = df[
    df["amazon_tipo_de_producto"].isna() |
    (df["amazon_tipo_de_producto"] == "")
].index.tolist()

print(f"🤖 LLM needed for {len(unmatched_indexes)} rows")

# Resume support
start_pos = 0
if os.path.exists(CHECKPOINT_FILE):
    with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
        start_pos = int(f.read().strip())
    print(f"🔁 Resuming LLM from batch index {start_pos}")

with Mistral(api_key=os.getenv("MISTRAL_API_TOKEN", "")) as mistral:

    for i in range(start_pos, len(unmatched_indexes), BATCH_SIZE):
        batch_indexes = unmatched_indexes[i:i + BATCH_SIZE]

        items = []
        for idx in batch_indexes:
            row = df.loc[idx]
            items.append({
                "seller_sku": row.get("seller-sku", ""),
                "title": row.get("item-name", ""),
                "brand": row.get("brand-name", ""),
                "description": row.get("item-description", ""),
                "bullet_points": [
                    row.get("bullet-point1", ""),
                    row.get("bullet-point2", ""),
                    row.get("bullet-point3", "")
                ]
            })

        guessed = guess_categories_batch(items, all_categories, mistral)

        for df_idx, category in zip(batch_indexes, guessed):
            df.at[df_idx, "amazon_tipo_de_producto"] = category

        # Atomic write
        df.to_csv(TMP_OUTPUT_FILE, index=False)
        os.replace(TMP_OUTPUT_FILE, OUTPUT_CSV)

        # Save checkpoint
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            f.write(str(i + BATCH_SIZE))

        print(f"✅ LLM classified rows {i + 1}–{min(i + BATCH_SIZE, len(unmatched_indexes))}")

        if stop_requested:
            print("💾 Progress safely saved. Exiting.")
            exit(0)

# Cleanup
if os.path.exists(CHECKPOINT_FILE):
    os.remove(CHECKPOINT_FILE)

print(f"🎉 Completed. Final output: {OUTPUT_CSV}")
