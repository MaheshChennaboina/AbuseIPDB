import requests
import pandas as pd
from datetime import datetime, timezone

def get_abuse_info(api_key, ip_address):
    url = f'https://api.abuseipdb.com/api/v2/check?ipAddress={ip_address}'
    headers = {
        'Key': api_key,
        'Accept': 'application/json'
    }

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Check if the request was successful

        data = response.json()
        abuse_confidence_score = data['data']['abuseConfidenceScore']
        abuse_reports = data['data'].get('reports', [])
        last_updated = data['data']['lastReportedAt']

        return abuse_confidence_score, len(abuse_reports), last_updated

    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")
    except Exception as err:
        print(f"An error occurred: {err}")

def calculate_time_elapsed(last_updated):
    if last_updated is None:
        return None, None, None

    last_updated_datetime = datetime.strptime(last_updated, "%Y-%m-%dT%H:%M:%S%z")
    
    # Make current_datetime timezone-aware
    current_datetime = datetime.utcnow().replace(tzinfo=timezone.utc)
    
    time_elapsed = current_datetime - last_updated_datetime

    # Calculate time elapsed in hours, weeks, and months
    time_elapsed_hours = time_elapsed.total_seconds() / 3600
    time_elapsed_weeks = time_elapsed.days / 7
    time_elapsed_months = time_elapsed.days / 30  # Approximation for months

    return time_elapsed_hours, time_elapsed_weeks, time_elapsed_months

def process_excel(input_file, api_key):
    # Read IPs from Excel input file
    df = pd.read_excel(input_file, names=['IP'])

    # Initialize output DataFrame
    output_data = {'IP': [], 'Abuse Confidence Score': [], 'Number of Reports': [], 'Last Updated': [],
                   'Time Elapsed (Hours)': [], 'Time Elapsed (Weeks)': [], 'Time Elapsed (Months)': []}

    for ip_address in df['IP']:
        abuse_info = get_abuse_info(api_key, ip_address)

        if abuse_info:
            abuse_confidence_score, totalReports, last_updated = abuse_info
            output_data['IP'].append(ip_address)
            output_data['Abuse Confidence Score'].append(abuse_confidence_score)
            output_data['Number of Reports'].append(totalReports)
            output_data['Last Updated'].append(last_updated)

            # Calculate time elapsed
            time_elapsed_hours, time_elapsed_weeks, time_elapsed_months = calculate_time_elapsed(last_updated)
            output_data['Time Elapsed (Hours)'].append(time_elapsed_hours)
            output_data['Time Elapsed (Weeks)'].append(time_elapsed_weeks)
            output_data['Time Elapsed (Months)'].append(time_elapsed_months)
        else:
            # If failed to retrieve abuse information, populate with NaN values
            output_data['IP'].append(ip_address)
            output_data['Abuse Confidence Score'].append(None)
            output_data['Number of Reports'].append(None)
            output_data['Last Updated'].append(None)
            output_data['Time Elapsed (Hours)'].append(None)
            output_data['Time Elapsed (Weeks)'].append(None)
            output_data['Time Elapsed (Months)'].append(None)

    # Create a DataFrame from the output data
    output_df = pd.DataFrame(output_data)

    # Generate the output file name with date and time
    current_datetime = datetime.now().strftime("%Y%m%d%H%M%S")
    output_file_name = f'output_{current_datetime}.xlsx'

    # Write the output DataFrame to the dynamically generated Excel file
    output_df.to_excel(output_file_name, index=False)

    print(f"Output file '{output_file_name}' created successfully.")

# Replace 'YOUR_API_KEY' with your actual AbuseIPDB API key
api_key = 'd245bd1979cbe344954820165650546329f8b2598a357927136ed319f9565fd5c5cb0bdcc9796ba6'

# Replace 'input.xlsx' with the name of your input Excel file
input_file_name = 'input.xlsx'

process_excel(input_file_name, api_key)
