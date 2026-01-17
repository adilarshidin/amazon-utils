from openpyxl import load_workbook
import pandas as pd
from pathlib import Path

excel_path = "templates/konus.xlsm"
sheet_name = "Plantilla"
csv_path = "input/konus_catalog.csv"
output_path = "output/amazon_konux.xlsm"

# Ensure output directory exists
Path(output_path).parent.mkdir(parents=True, exist_ok=True)

# Load Excel with macros preserved
wb = load_workbook(excel_path, keep_vba=True)
ws = wb[sheet_name]

# Read CSV safely (tolerant to malformed rows)
df = pd.read_csv(
    csv_path,
    encoding="latin-1",
    sep=";",
    engine="python",
    on_bad_lines="warn"
)

# Debug: read Excel row 4
row_4 = [cell.value for cell in ws[4]]
print("Excel - 4th row:")
print(row_4)

# Clear row 6 in Excel
for cell in ws[6]:
    cell.value = None

print("\nCSV - head():")
print(df.head(1))

# Save as NEW xlsm file
wb.save(output_path)

print(f"\nSaved modified file to: {output_path}")
