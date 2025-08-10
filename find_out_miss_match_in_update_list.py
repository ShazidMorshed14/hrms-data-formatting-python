import pandas as pd

# === Step 1: Read both Excel files ===
file1 = pd.read_excel("filtered_out_update_employee_list.xlsx")  # First file
file2 = pd.read_excel("hrms_employee_master_list_from_db_2025.xlsx")  # Second file

# === Step 2: Clean column names ===
file1.columns = file1.columns.str.strip()
file2.columns = file2.columns.str.strip()

# === Step 3: Merge on User Code ===
merged = file1.merge(
    file2,
    on="User Code",
    suffixes=("_file1", "_file2"),
    how="inner"  # only compare User Codes that exist in both
)

# === Step 4: Find mismatched Names ===
mismatches = merged[
    merged["Name_file1"].str.strip().str.lower() != merged["Name_file2"].str.strip().str.lower()
]

# === Step 5: Keep only relevant columns ===
mismatches = mismatches[["User Code", "Name_file1", "Name_file2"]]

# === Step 6: Save mismatches to a new Excel file ===
mismatches.to_excel("name_mismatches.xlsx", index=False)

print("✅ Name mismatches saved to 'name_mismatches.xlsx'")
