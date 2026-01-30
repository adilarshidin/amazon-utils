import pandas as pd

# Input and output file paths
input_file = "input/shopify_catalog_complete.xlsx"
output_file = "output/shopify_catalog_complete.csv"

# Read the Excel file
df = pd.read_excel(input_file)

df.to_csv(output_file, sep=',', index=False)

print(f"Converted '{input_file}' to '{output_file}' successfully!")
