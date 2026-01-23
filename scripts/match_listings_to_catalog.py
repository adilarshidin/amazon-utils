import pandas as pd

# Paths
ACTIVE_LISTINGS_PATH = "output/active_listings.csv"
CATALOG_PATH = "input/catalog_initial.csv"
OUTPUT_PATH = "output/matched_latest_catalog.csv"

# Required output columns (exact order)
CATALOG_COLUMNS = [
    "EAN",
    "NOMBRE",
    "COSTOS",
    "FECHA",
    "Fijo €",
    "Variable %",
    "Variable €",
    "Precio €",
    "Beneficio €",
    "Beneficio %",
    "PVP (Con Tax)",
    "Comision Amazon %",
    "Comision Amazon €",
    "Final PVP €",
]

# Load CSVs as strings to avoid numeric corruption
active_df = pd.read_csv(ACTIVE_LISTINGS_PATH, dtype=str)
catalog_df = pd.read_csv(CATALOG_PATH, dtype=str)

# Normalize keys
active_eans = (
    active_df["seller-sku"]
    .dropna()
    .astype(str)
    .str.strip()
)

catalog_eans = (
    catalog_df["EAN"]
    .dropna()
    .astype(str)
    .str.strip()
)

# Missing in catalog
missing_in_catalog = sorted(set(active_eans) - set(catalog_eans))

print(f"Active listings total: {len(set(active_eans))}")
print(f"Catalog EANs total: {len(set(catalog_eans))}")
print(f"Missing from catalog: {len(missing_in_catalog)}")

# Save for inspection
pd.DataFrame({"missing_seller_sku": missing_in_catalog}) \
  .to_csv("output/missing_from_catalog.csv", index=False)

# Normalize column names
active_df.columns = active_df.columns.str.strip()
catalog_df.columns = catalog_df.columns.str.strip()

# Validate required columns
if "seller-sku" not in active_df.columns:
    raise ValueError("active_listings.csv must contain 'seller-sku' column")

if "EAN" not in catalog_df.columns or "FECHA" not in catalog_df.columns:
    raise ValueError("catalog_initial.csv must contain 'EAN' and 'FECHA' columns")

# Parse FECHA with milliseconds support
catalog_df["FECHA"] = pd.to_datetime(
    catalog_df["FECHA"],
    format="%Y-%m-%d %H:%M:%S.%f",
    errors="coerce"
)

# Drop invalid rows
catalog_df = catalog_df.dropna(subset=["EAN", "FECHA"])

# Sort so latest FECHA is first
catalog_df = catalog_df.sort_values("FECHA", ascending=False)

# Deduplicate by EAN, keeping the latest FECHA
catalog_latest = catalog_df.drop_duplicates(subset=["EAN"], keep="first")

# Get active seller-sku values
active_eans = (
    active_df["seller-sku"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
)

# Match catalog entries
matched_catalog = catalog_latest[
    catalog_latest["EAN"].astype(str).isin(active_eans)
]

# Select required columns only
matched_catalog = matched_catalog[CATALOG_COLUMNS]

# Write output
matched_catalog.to_csv(OUTPUT_PATH, index=False)

print(f"Matched catalog written to: {OUTPUT_PATH}")
print(f"Total matched rows: {len(matched_catalog)}")
