import pandas as pd
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

INPUT_CSV = "output/all_listings_with_images_and_category.csv"
OUTPUT_CSV = "output/all_listings_with_images_and_category_translated.csv"
TMP_OUTPUT_FILE = "output/.all_listings_with_images_and_category_translated.tmp"
CHECKPOINT_FILE = "checkpoints/translate_product_type_checkpoint.txt"

BATCH_SIZE = 15
LLM_MODEL = "mistral-small-latest"

# =========================
# HELPERS
# =========================
def clean_json(text: str):
    """Remove code fences and parse JSON."""
    text = text.strip()
    text = re.sub(r"^```json\s*|\s*```$", "", text, flags=re.IGNORECASE)
    return json.loads(text)

def translate_product_types_batch(items, mistral):
    """
    Ask LLM to translate uppercase, underscore-separated product types
    into human-readable Spanish.
    """
    prompt = f"""
You are a professional translator. Translate each Amazon product type from uppercase
and underscore format to Spanish, in human-readable form. Keep the same order.
Return only a JSON array of strings, no explanations, no markdown.

Input product types:
{json.dumps(items, ensure_ascii=False)}

Output example:
["Producto A", "Producto B", ...]
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
df["amazon_product_type_es"] = df.get("amazon_product_type_es", "")

# =========================
# LLM TRANSLATION FOR UNTRANSLATED ROWS
# =========================
untranslated_indexes = df[
    df["amazon_product_type_es"].isna() |
    (df["amazon_product_type_es"] == "")
].index.tolist()

print(f"🤖 LLM needed for {len(untranslated_indexes)} rows")

# Resume support
start_pos = 0
if os.path.exists(CHECKPOINT_FILE):
    with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
        start_pos = int(f.read().strip())
    print(f"🔁 Resuming LLM from batch index {start_pos}")

with Mistral(api_key=os.getenv("MISTRAL_API_TOKEN", "")) as mistral:

    for i in range(start_pos, len(untranslated_indexes), BATCH_SIZE):
        batch_indexes = untranslated_indexes[i:i + BATCH_SIZE]

        items = [df.at[idx, "amazon_product_type"] for idx in batch_indexes]

        translated = translate_product_types_batch(items, mistral)

        for df_idx, translation in zip(batch_indexes, translated):
            df.at[df_idx, "amazon_product_type_es"] = translation

        # Atomic write
        df.to_csv(TMP_OUTPUT_FILE, index=False)
        os.replace(TMP_OUTPUT_FILE, OUTPUT_CSV)

        # Save checkpoint
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            f.write(str(i + BATCH_SIZE))

        print(f"✅ Translated rows {i + 1}–{min(i + BATCH_SIZE, len(untranslated_indexes))}")

        if stop_requested:
            print("💾 Progress safely saved. Exiting.")
            exit(0)

# Cleanup
if os.path.exists(CHECKPOINT_FILE):
    os.remove(CHECKPOINT_FILE)

print(f"🎉 Completed. Final output: {OUTPUT_CSV}")
