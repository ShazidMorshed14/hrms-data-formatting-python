import pandas as pd
import re

# Load your Excel file
df = pd.read_excel("final_updatable_employees.xlsx")
print('✅ Employee data read successfully!')

# Strip all column names
df.columns = df.columns.str.strip()

# Columns to clean
phone_columns = ['Mobile', 'Phone', 'Nominee Contact', 'MFS Account Number', 'Emergency Contact Mobile']

df['MFS Account Number'] = df['MFS Account Number'].apply(lambda x: str(x).replace('.0', '') if pd.notnull(x) else x)
df['Mobile'] = df['Mobile'].apply(lambda x: str(x).replace('.0', '') if pd.notnull(x) else x)
df['Emergency Contact Mobile'] = df['Emergency Contact Mobile'].apply(lambda x: str(x).replace('.0', '') if pd.notnull(x) else x)

# Function to format phone numbers into 11-digit 01xxxxxxxxx
def format_phone(number):
    
    if pd.isna(number):
        return ""
    
    # Convert to string and remove all non-digit characters
    number = str(number).strip()
    number = re.sub(r"\D", "", number)

    # Remove country code +880 or 880
    if number.startswith("880"):
        number = number[3:]
    elif number.startswith("+880"):
        number = number[4:]

    # If starts with 1 and length is 10, prepend 0
    if len(number) == 10 and number.startswith("1"):
        number = "0" + number

    # Ensure exactly 11 digits and starts with 01
    if len(number) == 11 and number.startswith("01"):
        return number
    else:
        return number
    

# Apply cleaning
for col in phone_columns:
    if col in df.columns:
        df[col] = df[col].apply(format_phone)
        print(f"✅ Cleaned phone column: {col}")
    else:
        print(f"⚠️ Column not found: {col}")

def safe_lower(val):
    return val.lower() if isinstance(val, str) else val

def safe_upper(val):
    return val.upper() if isinstance(val, str) else val



# # Normalize 'Emergency Contact Relationship'
# if 'Emergency Contact Relationship' in df.columns:
#     df['Emergency Contact Relationship'] = df['Emergency Contact Relationship'].astype(str).str.lower().replace({
#         'father': 'father',
#         'mother': 'mother',
#         'brother': 'brother',
#         'sister': 'sister',
#         'wife': 'spouse',
#         'husband': 'spouse',
#         'son': 'son',
#         'daughter': 'daughter',
#         'uncle': 'other',
#         'aunt': 'other',
#         '': ''
#     })
#     print('✅ Emergency Contact Relationship normalized')

# #Nominee Relation
# if 'Nominee Relation' in df.columns:
#     df['Nominee Relation'] = df['Nominee Relation'].astype(str).str.lower().replace({
#         'father': 'father',
#         'mother': 'mother',
#         'brother': 'brother',
#         'sister': 'sister',
#         'wife': 'spouse',
#         'husband': 'spouse',
#         'son': 'son',
#         'daughter': 'daughter',
#         'uncle': 'other',
#         'aunt': 'other',
#         '': ''
#     })
#     print('✅ Nominee Relation normalized')




#converting columns
df['MFS Type'] = df['MFS Type'].apply(safe_lower)
df['Gender'] = df['Gender'].apply(safe_lower)
df['Religion'] = df['Religion'].apply(safe_lower)
df['Marital Status'] = df['Marital Status'].apply(safe_lower)
df['Blood Group'] = df['Blood Group'].apply(safe_upper)
df['Status'] = df['Status'].astype(str).apply(safe_upper)
df['Emergency Contact Relationship'] = df['Emergency Contact Relationship'].apply(safe_lower)
df['Nominee Relation'] = df['Nominee Relation'].apply(safe_lower)



# Save to new Excel
df.to_excel("newly_formatted_updatable_employee_list.xlsx", index=False)
print("✅ File saved: formatted_new_employee_list.xlsx")