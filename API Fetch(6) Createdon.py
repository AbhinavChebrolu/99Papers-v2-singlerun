import requests
import pandas as pd
from datetime import datetime, timedelta
import time

# API endpoint and headers
url = "https://api.finanvo.in/company/newcompanies/list"
headers = {
    "Content-Type": "application/json",
    "x-api-key": "OliINWwpCj",
    "x-api-secret-key": "smdlme2Y9n25v8htM0uQkWGqxGdN3D8O8aLkf9Zl"
}

# Step 1: Fetch data from API for a specific date
def fetch_data_from_api(url, headers, date):
    try:
        params = {'type': 'date', 'bulkMasterType': 'CIN', 'date': date}
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data for {date}: {e}")
        return {}

# Step 2: Process and convert data to DataFrame
def process_data_to_dataframe(data):
    if 'cin' in data:
        companies = data['cin']
        company_list = []

        for company in companies:
            company_info = {
                "CIN": company.get("CIN", ""),
                "COMPANY_NAME": company.get("COMPANY_NAME", ""),
                "ACTIVITY_CODE": company.get("ACTIVITY_CODE", ""),
                "DATE_OF_REGISTRATION": company.get("DATE_OF_REGISTRATION", ""),
                "STATE": company.get("STATE", ""),
                "ROC": company.get("ROC", ""),
                "COMPANY_STATUS": company.get("COMPANY_STATUS", ""),
                "CATEGORY": company.get("CATEGORY", ""),
                "CLASS": company.get("CLASS", ""),
                "SUBCATEGORY": company.get("SUBCATEGORY", ""),
                "AUTHORIZED_CAPITAL": company.get("AUTHORIZED_CAPITAL", ""),
                "PAIDUP_CAPITAL": company.get("PAIDUP_CAPITAL", ""),
                "ACTIVITY_DESCRIPTION": company.get("ACTIVITY_DESCRIPTION", ""),
                "REGISTERED_OFFICE_ADDRESS": company.get("REGISTERED_OFFICE_ADDRESS", ""),
                "ADDRESS_OTHER_THAN_RO": company.get("ADDRESS_OTHER_THAN_RO", ""),
                "EMAIL": company.get("EMAIL", ""),
                "LAST_AGM": company.get("LAST_AGM", ""),
                "LAST_BALANCESHEET": company.get("LAST_BALANCESHEET", ""),
                "LISTING_STATUS": company.get("LISTING_STATUS", ""),
                "ACTIVE_COMPLIANCE": company.get("ACTIVE_COMPLIANCE", ""),
                "COMPANY_FILING_STATUS_16_17_18": company.get("COMPANY_FILING_STATUS_16_17_18", ""),
                "SUSPENDED_AT_STOCK": company.get("SUSPENDED_AT_STOCK", ""),
                "NUMBER_OF_MEMBERS": company.get("NUMBER_OF_MEMBERS", ""),
                "PARTNERS": company.get("PARTNERS", ""),
                "CIRP": company.get("CIRP", ""),
                "TIMESTAMP": company.get("TIMESTAMP", ""),
                "PINCODE": company.get("PINCODE", ""),
                "COUNTRY": company.get("COUNTRY", ""),
                "CITY": company.get("CITY", ""),
                "DISTRICT": company.get("DISTRICT", ""),
                "TOTAL_OBLIGATION_CONTRIBUTION": company.get("TOTAL_OBLIGATION_CONTRIBUTION", ""),
                "TYPE_OF_COMPANY": company.get("TYPE_OF_COMPANY", ""),
                # Capture the directors information
                "DIRECTORS": company.get("DIRECTORS", []),
                "CREATEDON": datetime.now(), # Add CREATEDON column with current timestamp
                "EmailMarketing1Status": "N" # Default to 'N' (not yet emailed)
            }
            company_list.append(company_info)

        return pd.DataFrame(company_list)
    else:
        print("No 'cin' data found in the response.")
        return pd.DataFrame()

# Step 3: Save the DataFrame to Excel, appending if CIN doesn't exist
def append_to_excel(df, filename="company_data1.xlsx"):
    try:
        try:
            existing_df = pd.read_excel(filename)
        except FileNotFoundError:
            existing_df = pd.DataFrame()

        if not df.empty:
            new_data_filtered = df[~df['CIN'].isin(existing_df['CIN'])]

            if not new_data_filtered.empty:
                if existing_df.empty:
                    merged_df = new_data_filtered
                else:
                    merged_df = pd.concat([existing_df, new_data_filtered]).drop_duplicates(subset='CIN', keep='last')
                
                merged_df.to_excel(filename, index=False)
                print(f"Data successfully appended to {filename}")
            else:
                print("No new data to append.")
        else:
            print("No new data to append.")

    except Exception as e:
        print(f"Error while saving to Excel: {e}")

# Step 4: Create a pivot table for count of `DATE_OF_REGISTRATION`
def create_pivot_table(filename="company_data1.xlsx"):
    try:
        df = pd.read_excel(filename)

        if not df.empty:
            pivot_df = df.groupby('DATE_OF_REGISTRATION').size().reset_index(name='COUNT')

            with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists="replace") as writer:
                pivot_df.to_excel(writer, sheet_name='companies_pivot', index=False)
                print(f"Pivot table successfully written to 'companies_pivot' sheet in {filename}")
        else:
            print("No data available to create pivot table.")
    
    except Exception as e:
        print(f"Error creating pivot table: {e}")

# Step 5: Fetch and save data until today's date
def fetch_and_save_data_until_today(start_date):
    current_date = datetime.strptime(start_date, "%Y-%m-%d")
    today = datetime.today()

    while current_date <= today:
        date_str = current_date.strftime("%Y-%m-%d")
        print(f"Fetching data for {date_str}...")

        raw_data = fetch_data_from_api(url, headers, date_str)

        if raw_data:
            df = process_data_to_dataframe(raw_data)
            append_to_excel(df, "company_data1.xlsx")

        current_date += timedelta(days=1)

# Main script
if __name__ == "__main__":
    start_date = "2024-12-24"

    while True:
        current_time = datetime.now()
        print(f"Current run started at: {current_time.strftime('%Y-%m-%d %H:%M:%S')}")

        fetch_and_save_data_until_today(start_date)

        finish_time = datetime.now()
        print(f"Current run finished at: {finish_time.strftime('%Y-%m-%d %H:%M:%S')}")

        next_run_time = finish_time + timedelta(minutes=5)
        print(f"Next run scheduled at: {next_run_time.strftime('%Y-%m-%d %H:%M:%S')}")

        create_pivot_table("company_data1.xlsx")

        print("Waiting for 5 minutes before next run...")
        time.sleep(300)  # Sleep for 5 minutes
