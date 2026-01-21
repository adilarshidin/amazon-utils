import pandas as pd
import asyncio
import random
import os
import re
import aiohttp
from pathlib import Path
from playwright.async_api import async_playwright, Page

# =========================
# FILES
# =========================
INPUT_CSV = "output/all_listings.csv"
OUTPUT_CSV = "output/all_listings_with_images.csv"
CHECKPOINT_CSV = "checkpoints/image_checkpoint.csv"
IMAGE_DIR = "downloaded_images"

Path(IMAGE_DIR).mkdir(parents=True, exist_ok=True)
Path(os.path.dirname(CHECKPOINT_CSV)).mkdir(parents=True, exist_ok=True)

# =========================
# SETTINGS
# =========================
CONCURRENCY = 3
IMAGE_WAIT_TIMEOUT = 15000

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_4_1) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.6 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
]

LANG_HEADERS = [
    {"Accept-Language": "es-ES,es;q=0.9,en;q=0.8"},
    {"Accept-Language": "en-US,en;q=0.9"},
    {"Accept-Language": "es-ES,es;q=0.9,en;q=0.8"},
]

# =========================
# LOAD INPUT
# =========================
df = pd.read_csv(INPUT_CSV)
assert "asin1" in df.columns, "asin1 column missing"

# =========================
# LOAD CHECKPOINT
# =========================
processed_asins = set()
if os.path.exists(CHECKPOINT_CSV):
    processed_asins = set(
        pd.read_csv(CHECKPOINT_CSV)["asin1"].astype(str).tolist()
    )

# =========================
# IMAGE HELPERS
# =========================
async def download_image(session: aiohttp.ClientSession, url: str, path: str):
    async with session.get(url) as resp:
        if resp.status == 200:
            with open(path, "wb") as f:
                f.write(await resp.read())

async def extract_landing_images(page: Page) -> list[str]:
    img = page.locator('img[data-a-image-name="landingImage"]')
    await img.wait_for(timeout=IMAGE_WAIT_TIMEOUT)

    dynamic = await img.get_attribute("data-a-dynamic-image")
    if not dynamic:
        src = await img.get_attribute("src")
        return [src] if src else []

    urls = re.findall(r'"(https://m\.media-amazon\.com[^"]+)"', dynamic)
    urls = {re.sub(r"\._[^.]+_", ".", u) for u in urls}
    return sorted(urls)

# =========================
# MAIN
# =========================
async def main():
    async with async_playwright() as p, aiohttp.ClientSession() as session:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()

        # PRE-CREATE PAGES (REUSED, NEVER CLOSED)
        pages: list[Page] = []
        for i in range(CONCURRENCY):
            page = await context.new_page()
            await page.set_extra_http_headers({
                **LANG_HEADERS[i % len(LANG_HEADERS)],
                "User-Agent": USER_AGENTS[i % len(USER_AGENTS)]
            })
            pages.append(page)

        semaphore = asyncio.Semaphore(CONCURRENCY)

        async def process_row(idx: int, asin: str):
            async with semaphore:
                if asin in processed_asins:
                    return

                page = pages[idx % CONCURRENCY]
                print(f"Processing ASIN {asin}")

                row_data = df.loc[idx].to_dict()

                try:
                    await page.goto(
                        f"https://www.amazon.es/dp/{asin}",
                        wait_until="domcontentloaded"
                    )

                    image_urls = await extract_landing_images(page)

                    for i, url in enumerate(image_urls, start=1):
                        filename = f"{asin}_image{i}.jpg"
                        filepath = os.path.join(IMAGE_DIR, filename)

                        await download_image(session, url, filepath)

                        row_data[f"image{i}"] = url
                        row_data[f"image{i}_file"] = filepath

                except Exception as e:
                    print(f"Error for ASIN {asin}: {e}")

                # WRITE OUTPUT ROW IMMEDIATELY
                pd.DataFrame([row_data]).to_csv(
                    OUTPUT_CSV,
                    mode="a",
                    header=not os.path.exists(OUTPUT_CSV),
                    index=False
                )

                # WRITE CHECKPOINT
                pd.DataFrame([{"asin1": asin}]).to_csv(
                    CHECKPOINT_CSV,
                    mode="a",
                    header=not os.path.exists(CHECKPOINT_CSV),
                    index=False
                )

                processed_asins.add(asin)
                await asyncio.sleep(random.uniform(2, 4))

        tasks = [
            process_row(idx, str(row["asin1"]))
            for idx, row in df.iterrows()
            if pd.notna(row["asin1"])
        ]

        await asyncio.gather(*tasks)

        # BROWSER INTENTIONALLY LEFT OPEN
        print("Finished processing all ASINs. Browser left open.")

# =========================
# RUN
# =========================
asyncio.run(main())
