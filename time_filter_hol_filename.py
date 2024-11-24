import pandas as pd
from datetime import datetime, timedelta
from pandas.tseries.holiday import USFederalHolidayCalendar

# Function to calculate duration between 6:00 AM and 10:00 PM, excluding weekends and federal holidays
def calculate_filtered_duration(start, end, holidays):
    if pd.notnull(start) and pd.notnull(end):
        # Ensure start and end are datetime
        start = pd.to_datetime(start)
        end = pd.to_datetime(end)
        
        # Initialize total duration
        total_duration = pd.Timedelta(0)
        
        current = start
        while current < end:
            # Skip weekends and federal holidays
            if current.weekday() < 5 and current.date() not in holidays:  # Weekday: 0=Monday, 4=Friday
                day_start = current.replace(hour=6, minute=0, second=0)
                day_end = current.replace(hour=22, minute=0, second=0)
                
                # Calculate valid range within the day
                valid_start = max(current, day_start)
                valid_end = min(end, day_end)
                
                if valid_start < valid_end:
                    total_duration += (valid_end - valid_start)
            
            # Move to the next day
            current += pd.Timedelta(days=1)
            current = current.replace(hour=0, minute=0, second=0)
        
        # Convert duration to days, hours, minutes
        days = total_duration.days
        hours, remainder = divmod(total_duration.seconds, 3600)
        minutes = remainder // 60
        return f"{days} days, {hours} hours, {minutes} minutes"
    
    return None

# Prompt user for year and filename
year = int(input("Enter the year to calculate federal holidays: "))
file_path = input("Enter the filename (with extension) of the file to process: ")

# Generate the list of federal holidays for the given year
cal = USFederalHolidayCalendar()
holidays = cal.holidays(start=f"{year}-01-01", end=f"{year}-12-31").to_pydatetime()
holidays = {holiday.date() for holiday in holidays}

# Load the dataset
try:
    df = pd.read_excel(file_path)
except FileNotFoundError:
    print(f"File '{file_path}' not found. Please ensure the filename and path are correct.")
    exit()

# Ensure your columns match these names
start_col = 'Start Date and Time'
end_col = 'End Date and Time'

# Check if the required columns exist
if start_col not in df.columns or end_col not in df.columns:
    print(f"The file must contain '{start_col}' and '{end_col}' columns.")
    exit()

# Apply the function to each row
df['Total Record'] = df.apply(
    lambda row: calculate_filtered_duration(row[start_col], row[end_col], holidays),
    axis=1
)

# Save the updated dataset
output_path = f"filtered_output_with_holidays_{year}.xlsx"
df.to_excel(output_path, index=False)

print(f"Filtered data saved to {output_path}")
