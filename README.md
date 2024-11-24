
# Incident time filter calculation 
<i>(removes time outside business hours)</i>

This is a python script desined to output an xlsx (Excel) file with total incident duration that excludes time outside set business hours.
In this example, the script uses 6:00 am through 10:00 pm range Monday through Friday as business hours and will exclude any time before or after that range.
It will also exclude US Federal Holidays and Weekends.
There will be two inputs asked from the script.
1. the year for the incident which dates are set to so it imports the correct US Holiday schedule
2. the filename - this must be a file with .xlsx extension

This helps using a local (offline) way to calculate incident/event duration without providing identifiable information outside of the working environment. All that is needed is start date/time and end date/time plus a filename you created for this calculation and the year. Ensure these columns contain both date and time (format them accordingly within Excel or LibreOffice Calc)

# Installation
dependencies
ensure python3-pip is installed:

sudo apt install python3-pip

then install pandas and openpyxl:

pip install pandas openpyxl

if installation says scripts not in path, add to path like this: export PATH=$PATH:/path/to/directory

# <b>Usage:</b>

1. create an excel.xlsx file where column A is for example a record number, column B is the start date and time, column C is the end date and time and column D is total record (total time of incident). See example included. what you name each column is up to you but if you change it, you need to adjust/mod the script to accomodate that change. those are under two sections
     Ensure your column matches these names <--- this is where you enter the start and end date/time column names
     Apply the function to each row <--- this is where you enter the total time result

3. run the script as such: python3 time_filter_hol_filename.py (note the script can or rather should be in the path of the file you created to make it easier).
4. the script will ask you for the year so it determines the right US Holidays and also for the filename where if file is called test.xlsx you enter test.xlsx.
5. the script will generate a file named filtered_output_with_holidays_<year>.xlsx (e.g., filtered_output_with_holidays_2024.xlsx)

EOF
