import pandas as pd

# === Step 1: Read both Excel files ===
main_df = pd.read_excel("filterted_new_employee_list.xlsx")   # File with Roles and empty role_id
roles_df = pd.read_excel("roles_list.xlsx")  # File with id and role

# Ensure consistent column names (strip spaces)
main_df.columns = main_df.columns.str.strip()
roles_df.columns = roles_df.columns.str.strip()

print("main file cols", main_df.columns)
print("role file cols", roles_df.columns)

# Normalize text for matching
main_df['Roles'] = main_df['Roles'].astype(str).str.strip().str.lower()
roles_df['role'] = roles_df['role'].astype(str).str.strip().str.lower()

# === Step 2: Merge to get role_id ===
merged_df = main_df.merge(
    roles_df[['id', 'role']], 
    left_on='Roles', 
    right_on='role', 
    how='left'
)

# Fill role_id with the id from roles_df
merged_df['role_id'] = merged_df['id']

# === Step 3: Keep only Roles and role_id ===
final_df = merged_df[['Roles', 'role_id']]

# === Step 4: Save updated file ===
final_df.to_excel("main_file_with_role_ids.xlsx", index=False)

print("✅ File saved with only 'Roles' and 'role_id'")
