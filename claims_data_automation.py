"""
Claims Data Consolidation & Deduplication Automation
"""

import pandas as pd
import os

print(os.getcwd())

folder_path = r"data/input"

final_df = pd.DataFrame()

for file in os.listdir(folder_path):

    if file.endswith(".xlsx"):

        file_path = os.path.join(folder_path, file)

        summary = pd.read_excel(file_path, header=None, nrows=8)

        insured = summary.iloc[7, 0]
        policy = summary.iloc[7, 3]
        valuation_date = summary.iloc[7, 5]

        table = pd.read_excel(file_path, skiprows=9)

        table.columns = table.columns.str.strip()

        claim_number_cols = [
            col for col in table.columns if "CLAIM NUMBER" in col.upper()
        ]

        if not claim_number_cols:
            print(f"Skipped (no claim table): {file}")
            continue

        claim_number_col = claim_number_cols[0]

        table = table[table[claim_number_col].notna()]

        if table.empty:
            print(f"Skipped (no claim rows): {file}")
            continue

        table.insert(0, "Insured Name", insured)
        table.insert(1, "Policy Number", policy)
        table.insert(2, "Valuation Date", valuation_date)

        final_df = pd.concat([final_df, table], ignore_index=True)

before_count = len(final_df)

final_df = final_df.drop_duplicates(subset="CLAIM NUMBER", keep="first")

after_count = len(final_df)

print(f"Duplicates removed: {before_count - after_count}")

output_folder = "data/output"
os.makedirs(output_folder, exist_ok=True)

output_path = os.path.join(output_folder, "final_output.xlsx")
final_df.to_excel(output_path, index=False)

print("Consolidated file created successfully.")
