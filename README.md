# Amazon Utils

## Features

1. Product names translator using Mistral LLM API.

## Setup

1. Provide env variables.
2. Provide constants in the necessary script in `scripts/`.
3. Provide files to work with in `input` directory of the project root.

## Pipelines

1. Filtering spreadsheets:

    1) EAN-ASIN conversion. Check US and ES markets.

    2) Product names translation.

    3) Deduplicate the products with the same ASIN/EAN codes leaving only the most recent ones.

    4) Scrape Amazon and remove the products that are not available or can not be imported.

2. Amazon listing prices update:

    1) Download price & quantity template. Or use Flat.File.PriceInventory.es.xlsx from templates directory.

    2) Fill sku column (EAN or other seller side code).

    3) Fill price and quantity (omit any if not needed to change).

    3) Upload via "Add product" in Inventory Management section.

3. Amazon listing delete guide.

    1) In searchbar search "inventory loader template". Or use Flat.File.InventoryLoader.us.xlsx from templates directory and skip to step 5.

    2) Switch country to US.

    3) Click on "Use the inventory loader" article.

    4) Download template.

    5) Fill sku column (EAN or other seller side code).

    6) Fill add-delete column (d for delete from shop but keep in market, x to delete entirely)

### Tools needed

1. Scraper of item prices sold by other shops to find the reasonable margin.
2. Seller Assistant like info for finding best selling items and other insights.
3. EAN scraper to find on which Amazon market the item exists and is available.

### Tasks

1. Form full catalog uniting all products from sellerboard with existing large catalog.
Keep only the latest duplicate. Check that the format of ALL columns is the same.
Money columns are formatted correctly, percentage columns too.
