import pandas as pd
from datetime import datetime

# Load the Excel file
excel_file = "mail_logs.xlsx"  # Change path if needed
sheet_name = "Sheet1"          # Adjust sheet name if different

# Step 1: Ask user for input dates
date_input = input("Enter date(s) in YYYY-MM-DD format separated by commas: ")
date_list = [date.strip() for date in date_input.split(",")]

try:
    # Step 2: Read Excel and convert timestamp to datetime
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    df['Timestamp'] = pd.to_datetime(df['Timestamp'])
    df['Date'] = df['Timestamp'].dt.date.astype(str)  # Extract just the date in string

    # Step 3: Filter for input dates
    filtered_df = df[df['Date'].isin(date_list)]

    # Step 4: Count emails by sender per date
    report = filtered_df.groupby(['Date', 'Sender Email']).size().reset_index(name='Email Count')

    # Step 5: Save the report
    output_file = "email_report.xlsx"
    report.to_excel(output_file, index=False)
    print(f"\nReport generated successfully and saved to {output_file}.")

except Exception as e:
    print(f"An error occurred: {e}")
