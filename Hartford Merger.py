# Importing all the required libraries
import pandas as pd
import os
print(os.getcwd())

# Defining the folder path
folder_path = r"C:\Users\Anam\OneDrive - RG Insurance Processing Services Private Limited\Work\Automated pdf conversion\Hardford Merger\Testing\Hartford Merger\Hartford LR's"

# Creating a DataFrame
final_df = pd.DataFrame()

# Looping through all files in the folder
for file in os.listdir(folder_path):

    if file.endswith(".csv"):

        file_path = os.path.join(folder_path, file)

        # Fetching the data from the summary
        summary = pd.read_excel(file_path, header=None, nrows=8)

        insured = summary.iloc[7, 0]
        policy = summary.iloc[7, 3]
        valuation_date = summary.iloc[7, 5]

        # Fetching the table data
        table = pd.read_excel(file_path, skiprows=9)

        # Clean column names
        table.columns = table.columns.str.strip()

        # Detect the real claim number column specifically
        claim_number_cols = [col for col in table.columns if "CLAIM NUMBER" in col.upper()]

        if not claim_number_cols:
            print(f"Skipped (no claim table): {file}")
            continue

        claim_number_col = claim_number_cols[0]

        # Keep only real claim rows
        table = table[table[claim_number_col].notna()]

        # Skip files with no claim data
        if table.empty:
           print(f"skipped (no claim rows): {file}")
           continue

        # Inserting summary columns
        table.insert(0, "Insured Name", insured)
        table.insert(1, "Policy Number", policy)
        table.insert(2, "Valuation Date", valuation_date)

        # Appending into final dataframe
        final_df = pd.concat([final_df, table], ignore_index=True)

# Eliminating duplicate values
before_count = len(final_df)

final_df = final_df.drop_duplicates(subset="CLAIM NUMBER", keep="first")

after_count = len(final_df)

print(f"Duplicates removed: {before_count - after_count}")

# Creating the output file
output_path = os.path.join(folder_path, "AW Hartford.xlsx")
final_df.to_excel(output_path, index=False)