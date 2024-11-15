import pandas as pd
import re

# Load the uploaded file
file_path = ''

# Load the three attendance sheets
f23_attendance = pd.read_excel(file_path, sheet_name='F23 Attendance')
s24_attendance = pd.read_excel(file_path, sheet_name='S24 Attendance')
f24_attendance = pd.read_excel(file_path, sheet_name='F24 Attendance')

# Function to clean and prepare the data
def clean_and_prepare_data(df):
    # Standardize header names and remove extra spaces
    df.columns = df.columns.str.strip().str.lower()

    # Drop rows that are repeated headers or have empty timestamps
    df = df[~df['timestamp'].astype(str).str.contains("timestamp", case=False)]
    df = df.dropna(subset=['timestamp'])

    # Function to remove email addresses from text
    def remove_emails(text):
        return re.sub(r'\b\S+@\S+\.\S+\b', '', str(text)).strip()

    full_names = []
    for index, row in df.iterrows():
        first_name, last_name = '', ''

        # Detect where the first and last name columns are located
        if 'first name' in df.columns and 'last name' in df.columns:
            first_name = row['first name'] if not pd.isna(row['first name']) else ''
            last_name = row['last name'] if not pd.isna(row['last name']) else ''
        elif 'student email' in df.columns and 'first name' in df.columns:
            # Handle shifts: 'Student Email' followed by 'First Name'
            first_name = row['first name'] if not pd.isna(row['first name']) else ''
            last_name = row['student email'] if not pd.isna(row['student email']) else ''
        elif 'first name' in df.columns and 'student email' in df.columns:
            # Handle another shift
            first_name = row['student email'] if not pd.isna(row['student email']) else ''
            last_name = row['first name'] if not pd.isna(row['first name']) else ''
        else:
            continue  # Skip if columns are not recognizable

        # Remove emails from both names and combine
        full_name = f"{remove_emails(first_name)} {remove_emails(last_name)}".strip()
        full_names.append(full_name)

    # Assign the cleaned full names
    df['Full Name'] = full_names

    # Drop rows where 'Full Name' is empty or 'nan'
    df = df[['timestamp', 'Full Name']].dropna()
    df = df[df['Full Name'].str.strip() != '']

    return df

# Apply cleaning to each sheet
f23_cleaned = clean_and_prepare_data(f23_attendance)
s24_cleaned = clean_and_prepare_data(s24_attendance)
f24_cleaned = clean_and_prepare_data(f24_attendance)

# Combine all cleaned data
combined_data = pd.concat([f23_cleaned, s24_cleaned, f24_cleaned], ignore_index=True)

# Save the combined cleaned data to a new Excel file
combined_data.to_excel('', index=False)

