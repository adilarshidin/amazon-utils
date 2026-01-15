import csv

# === File paths (constants) ===
INPUT_CSV_PATH = "output/all_listings_duplicates_deleted.csv"
OUTPUT_TSV_PATH = "output/all_listings_duplicates_deleted.tsv"


def csv_to_tsv(input_path: str, output_path: str) -> None:
    with open(input_path, mode="r", newline="", encoding="utf-8") as csv_file, \
         open(output_path, mode="w", newline="", encoding="utf-8") as tsv_file:

        csv_reader = csv.reader(csv_file)
        tsv_writer = csv.writer(tsv_file, delimiter="\t")

        for row in csv_reader:
            tsv_writer.writerow(row)


if __name__ == "__main__":
    csv_to_tsv(INPUT_CSV_PATH, OUTPUT_TSV_PATH)
