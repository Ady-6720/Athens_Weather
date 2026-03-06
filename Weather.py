import pandas as pd
import requests
import datetime
import os

# 1. Configuration
SAVE_DIR = r"C:\UGA_Weather"
MASTER_FILE = os.path.join(SAVE_DIR, "KAHN_Weather_Master.xlsx")
API_URL = "https://api.weather.com/v1/location/KAHN:9:US/observations/historical.json"
API_KEY = "PASTE YOUR API KEY HERE"

if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)

# 2. Dynamic Date Calculation (Last 31 Days)
end_dt = datetime.datetime.now() - datetime.timedelta(days=1)
start_dt = end_dt - datetime.timedelta(days=30)

start_str = start_dt.strftime('%Y%m%d')
end_str = end_dt.strftime('%Y%m%d')

print(f"Dynamic window: {start_str} to {end_str}")

# 3. Fetch Data
params = {
    "apiKey": API_KEY,
    "units": "e",
    "startDate": start_str,
    "endDate": end_str
}

print("Fetching data from API...")
try:
    response = requests.get(API_URL, params=params)
    response.raise_for_status()
    data = response.json()
    new_raw = pd.DataFrame(data.get('observations', []))
except Exception as e:
    print(f"Error fetching data: {e}")
    exit()

if new_raw.empty:
    print("No data returned for this period.")
    exit()

# 4. Processing
# Time conversion and rounding to hour
new_raw['observations.valid_time_est'] = (
    pd.to_datetime(new_raw['valid_time_gmt'], unit='s', utc=True)
    .dt.tz_convert('US/Eastern')
    .dt.floor('h')
)

new_raw = new_raw.drop_duplicates(subset=['observations.valid_time_est'])
new_raw = new_raw.set_index('observations.valid_time_est').sort_index()

# Reindex and Interpolate
full_range = pd.date_range(start=new_raw.index.min(), end=new_raw.index.max(), freq='h')
new_raw = new_raw.reindex(full_range)

cols_to_fix = ['temp', 'dewPt', 'rh', 'pressure', 'wspd', 'precip_hrly']
new_raw[cols_to_fix] = new_raw[cols_to_fix].interpolate(method='linear', limit=2)

# Formatting
new_raw = new_raw.reset_index().rename(columns={'index': 'Timestamp'})
new_raw['Date'] = new_raw['Timestamp'].dt.strftime('%Y-%m-%d')
new_raw['Day of Week'] = new_raw['Timestamp'].dt.day_name()
new_raw['Time'] = new_raw['Timestamp'].dt.strftime('%I:%M %p')

# Rounding
new_raw[['temp', 'dewPt', 'rh', 'wspd']] = new_raw[['temp', 'dewPt', 'rh', 'wspd']].round(0)
new_raw[['pressure', 'precip_hrly']] = new_raw[['pressure', 'precip_hrly']].round(2)

column_mapping = {
    'Date': 'Date', 'Day of Week': 'Day of Week', 'Time': 'Time',
    'temp': 'Temperature (°F)', 'dewPt': 'Dew Point (°F)',
    'rh': 'Relative Humidity (%)', 'pressure': 'Pressure (in.)',
    'wspd': 'Wind Speed (mph)', 'precip_hrly': 'Hourly Precipitation (in.)'
}

new_final = new_raw[list(column_mapping.keys())].rename(columns=column_mapping)

# 5. Merge with Master File
if os.path.exists(MASTER_FILE):
    existing_df = pd.read_excel(MASTER_FILE)
    final_df = pd.concat([existing_df, new_final], ignore_index=True).drop_duplicates(subset=['Date', 'Time'])
    print("Appended new data to existing master file.")
else:
    final_df = new_final
    print("Created new master file.")

# Final Save
final_df.fillna(" ").to_excel(MASTER_FILE, index=False)
print(f"Success! Master file updated at: {os.path.abspath(MASTER_FILE)}")
