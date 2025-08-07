import pandas as pd
import re

# Load your Excel file
df = pd.read_excel("new_emplyee_requests.xlsx")
print('✅ Employee data read successfully!')

# Strip all column names to avoid mismatch due to extra spaces
df.columns = df.columns.str.strip()

# Columns to clean
phone_columns = ['Mobile', 'Phone', 'Nominee Contact', 'MFS Account Number', 'Emergency Contact Mobile']



# === Function to clean and format phone numbers ===
def format_phone(number):
    print(f"🔍 Formatting phone number: {number}")
    if pd.isna(number):
        return ""
    number = str(number)
    number = str(number).strip()
    
    # Remove all non-numeric characters
    number = re.sub(r"\D", "", number)

    # If number starts with 8801 and is 13 digits, strip to last 11 digits
    if len(number) == 11 & number.startswith("01"):
        pass
    elif len(number) > 11 & number.startswith("88"):
        number = number[-11:]
    elif len(number) > 11 & number.startswith("+88"):
        number = number[-11:]    
    elif len(number)==10 & number.startswith("1"):
        number = "0" + number
    else:
        print(f"⚠️ Invalid phone number format: {number}")
        number = "invalid"  # Invalid format

    return number

# === Apply cleaning to all phone number columns ===
for col in phone_columns:
    if col in df.columns:
        df[col] = df[col].apply(format_phone)
        print(f"✅ Cleaned phone column: {col}")
    else:
        print(f"⚠️ Column not found in Excel: {col}")

# === Convert other columns to lowercase or normalize ===

# Convert 'MFS Type' to lowercase
if 'MFS Type' in df.columns:
    df['MFS Type'] = df['MFS Type'].astype(str).str.lower()
    print('✅ MFS Type converted to lowercase')

# Normalize 'Emergency Contact Relationship'
if 'Emergency Contact Relationship' in df.columns:
    df['Emergency Contact Relationship'] = df['Emergency Contact Relationship'].astype(str).str.lower().replace({
        'father': 'father',
        'mother': 'mother',
        'brother': 'brother',
        'sister': 'sister',
        'wife': 'spouse',
        'husband': 'spouse',
        'son': 'son',
        'daughter': 'daughter',
        'uncle': 'other',
        'aunt': 'other',
         
    })
    print('✅ Emergency Contact Relationship normalized')

# Gender
if 'Gender' in df.columns:
    df['Gender'] = df['Gender'].astype(str).str.lower()
    print('✅ Gender converted to lowercase')

# Religion
if 'Religion' in df.columns:
    df['Religion'] = df['Religion'].astype(str).str.lower()
    print('✅ Religion converted to lowercase')

# Marital Status
if 'Marital Status' in df.columns:
    df['Marital Status'] = df['Marital Status'].astype(str).str.lower()
    print('✅ Marital Status converted to lowercase')

# === Save to new Excel file ===
df.to_excel("formatted_new_employee_list.xlsx", index=False)
print("✅ File saved: formatted_new_employee_list.xlsx")
