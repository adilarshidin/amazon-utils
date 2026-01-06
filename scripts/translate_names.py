import os
import csv
import re
import json
import signal
from dotenv import load_dotenv
from mistralai import Mistral

load_dotenv()

INPUT_FILE = "output/asin_results.csv"
OUTPUT_FILE = "output/translated_catalog.csv"
CHECKPOINT_FILE = "input/translate_checkpoint.txt"
TMP_OUTPUT_FILE = "output/.translated_catalog.tmp"

BATCH_SIZE = 20


def clean_json(text: str):
    text = text.strip()
    text = re.sub(r"^```json\s*|\s*```$", "", text, flags=re.IGNORECASE)
    return json.loads(text)


def translate_batch(names, mistral):
    prompt = f"""
You are given a list of product names from an Amazon catalog.

Rules:
- Translate each name to English.
- Use the product context to choose the most accurate translation.
- Return ONLY a JSON array of strings in the same order.
- No markdown, no explanations.

Product names:
{names}
"""
    res = mistral.chat.complete(
        model="mistral-small-latest",
        messages=[{"role": "user", "content": prompt}],
        stream=False,
    )

    return clean_json(res.choices[0].message.content)


def load_rows():
    if os.path.exists(OUTPUT_FILE):
        print("🔁 Loading partial output file...")
        with open(OUTPUT_FILE, newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))
    else:
        with open(INPUT_FILE, newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))


# Graceful Ctrl+C handler
stop_requested = False


def handle_sigint(signum, frame):
    global stop_requested
    stop_requested = True
    print("\n🛑 Ctrl+C detected. Finishing current batch safely...")


signal.signal(signal.SIGINT, handle_sigint)


rows = load_rows()

start_index = 0
if os.path.exists(CHECKPOINT_FILE):
    with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
        start_index = int(f.read().strip())
    print(f"🔁 Resuming from row {start_index + 1}")
else:
    print("📄 Starting fresh translation...")


with Mistral(api_key=os.getenv("MISTRAL_API_KEY", "")) as mistral:

    for i in range(start_index, len(rows), BATCH_SIZE):
        batch = rows[i : i + BATCH_SIZE]
        names = [row["NOMBRE"] for row in batch]

        translated = translate_batch(names, mistral)

        for row, new_name in zip(batch, translated):
            row["NOMBRE"] = new_name

        # Atomic write
        with open(TMP_OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=rows[0].keys())
            writer.writeheader()
            writer.writerows(rows)

        os.replace(TMP_OUTPUT_FILE, OUTPUT_FILE)

        # Save checkpoint
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            f.write(str(i + BATCH_SIZE))

        print(f"✅ Saved rows {i + 1}–{min(i + BATCH_SIZE, len(rows))}")

        if stop_requested:
            print("💾 Progress safely saved. Exiting.")
            exit(0)


# Cleanup on full completion
os.remove(CHECKPOINT_FILE)
print(f"🎉 Translation complete. Saved to {OUTPUT_FILE}")
