import pandas as pd
from datetime import datetime

# Path to your CSV file
file_path = "bulk-employee-update-template-till-done.csv"

# Columns to validate (change these to match your CSV headers)
date_columns = ["joining_date", "confirmation_date", "date_of_birth"]

# Function to check if a date is valid and in YYYY-MM-DD format
def is_valid_date(date_str):
    try:
        datetime.strptime(str(date_str), "%Y-%m-%d")
        return True
    except ValueError:
        return False

# Read CSV
df = pd.read_csv(file_path)

# Find rows where at least one date column is invalid
invalid_rows = df[~df[date_columns].applymap(is_valid_date).all(axis=1)]

# Save invalid rows to a new CSV for review
invalid_rows.to_csv("invalid_dates.csv", index=False)

print(f"Found {len(invalid_rows)} rows with invalid dates. Saved to 'invalid_dates.csv'.")
