import pandas as pd

# === Step 1: Read files ===
main_df = pd.read_excel("filterted_new_employee_list.xlsx")  # Main file
points_df = pd.read_excel("point_by_distributor.xlsx")       # Reference file

# Strip spaces in column names
main_df.columns = main_df.columns.str.strip()
points_df.columns = points_df.columns.str.strip()

print ("Main file columns:", main_df.columns)
print ("Points file columns:", points_df.columns)

# Store original case for output
main_df['dist_orig'] = main_df['Distribution House']
main_df['terr_orig'] = main_df['Territory']
main_df['point_orig'] = main_df['Distributor Points']

# Normalize for matching
def normalize(series):
    return series.astype(str).str.strip().str.lower()

main_df['Distribution House'] = normalize(main_df['Distribution House'])
main_df['Territory'] = normalize(main_df['Territory'])
main_df['Distributor Points'] = normalize(main_df['Distributor Points'])

points_df['distributor'] = normalize(points_df['distributor'])
points_df['territory'] = normalize(points_df['territory'])
points_df['point'] = normalize(points_df['point'])

# === Step 2: Merge ===
merged_df = main_df.merge(
    points_df[['point_id', 'distributor', 'territory', 'point']],
    left_on=['Distribution House', 'Territory', 'Distributor Points'],
    right_on=['distributor', 'territory', 'point'],
    how='left'
)

# === Step 3: Keep only desired output columns ===
output_df = merged_df[['dist_orig', 'terr_orig', 'point_orig', 'point_id']]

# Rename to final names
output_df.columns = ['distributor', 'territory', 'point', 'point_id']

# === Step 4: Save output ===
output_df.to_excel("mapped_points.xlsx", index=False)

# Debug info
matches_found = output_df['point_id'].notna().sum()
print(f"✅ Mapped points saved to 'mapped_points.xlsx' — {matches_found} matches found.")
