import pandas as pd
import re
from datetime import datetime

def format_part(part):
    # For parts greater than or equal to 10, split into two numbers
    if len(part) > 2:
        return str(int(part[0:2])) + '/' + str(int(part[2:]))
    # For parts of two digits, split similarly
    elif len(part) == 2:
        return str(int(part[0])) + '/' + str(int(part[1]))
    # For single digit parts, handle them separately
    return str(int(part[0])) + '/' + '2'   # Assuming '2' as a default for single digits

def fix_title(title):
    # Split the title by '-'
    parts = title.split('-')
    
    # Apply format_part to each part
    part1 = format_part(parts[0])
    part2 = format_part(parts[1])

    # Join the parts with '-'
    fixed_title = part1 + '-' + part2
    return fixed_title

def read_xls(xls_path):
    # Read the Excel file into a dictionary of DataFrames, where the key is the tab name
    xls = pd.ExcelFile(xls_path)
    
    # Read visible sheets
    data_frames = {}
    for sheet in xls.book.worksheets:
        if sheet.sheet_state == "visible":
            print(f"sheet:{sheet.title} is {sheet.sheet_state}")
            try:
                data_frames[fix_title(sheet.title)] = pd.read_excel(xls_path, sheet_name=sheet.title, header=None, index_col=None)
                print(f"Sheet saved as data frame with title: {fix_title(sheet.title)}")
            except:
                print("Title was not of correct format, not saving")
    return data_frames

def find_sunday_position(df):
    # Search for the cell containing "Sunday"
    for row_idx in range(df.shape[0]):
        for col_idx in range(df.shape[1]):
            if str(df.iloc[row_idx, col_idx]).lower() == "sunday":
                return row_idx, col_idx
    return None, None  # If "Sunday" isn't found

def fix_date(date_str:str):
    # Get current month and year
    today = datetime.today()
    current_month = today.month
    current_year = today.year

    # Mapping of month names to numbers
    month_map = {
        "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
        "May": "05", "Jun": "06", "Jul": "07", "Aug": "08",
        "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12"
    }

    # Split the date input, assuming format like "26th Jan"
    parts = date_str.split()
    day = ''.join(filter(str.isdigit, parts[0]))  # Extract numeric part of day
    month_name = parts[1]  # Extract month name

    # Convert month name to numeric format
    if month_name not in month_map:
        raise ValueError(f"Invalid month name in date: {date_str}")

    month_number = month_map[month_name]

    # Determine the correct year based on the rules
    if current_month in [1, 2, 3] and month_number == "12":
        year = current_year - 1  # December belongs to the previous year
    elif current_month in [10, 11, 12] and month_number == "01":
        year = current_year + 1  # January belongs to the next year
    else:
        year = current_year  # Otherwise, it's the current year

    # Format as DD/MM/YYYY
    formatted_date = f"{day.zfill(2)}/{month_number}/{year}"
    
    return formatted_date
    
def extract_shift_details(shift_time):
    # Match HH:MM-HH:MM format
    match = re.match(r"(\d{1,2}:\d{2})-(\d{1,2}:\d{2})", shift_time)
    
    if not match:
        raise ValueError(f"Invalid shift format: {shift_time}")
    
    start_time_str, end_time_str = match.groups()
    
    # Convert to datetime objects for time calculations
    time_format = "%H:%M"
    start_time = datetime.strptime(start_time_str, time_format)
    end_time = datetime.strptime(end_time_str, time_format)

    # Calculate duration in hours
    duration = (end_time - start_time).total_seconds() / 3600

    return start_time.strftime("%H:%M"), duration

def find_shifts(data_frames, name):
    results = []

    # Loop through each sheet (data frame)
    for sheet_title, df in data_frames.items():
        sunday_row, sunday_col = find_sunday_position(df)

        # If "Sunday" isn't found, skip the sheet
        if sunday_row is None or sunday_col is None:
            continue

        # Dates are one row down from Sunday, and shift times are one column to the left of Sunday
        shift_time_row = sunday_row + 1
        shift_col = sunday_col - 1
        date_row = sunday_row + 1
        date_col_start = sunday_col

        # Iterate through the dates and shift times
        for col_idx in range(date_col_start, df.shape[1]):  # Loop through dates
            date = df.iloc[date_row, col_idx]
            for row_idx in range(shift_time_row, df.shape[0]):  # Loop through shift times
                shift_time = df.iloc[row_idx, shift_col]
                if pd.notna(df.iloc[row_idx, col_idx]) and name.lower() in df.iloc[row_idx, col_idx].lower():
                    start_time, duration = extract_shift_details(shift_time)
                    results.append({"Sheet": sheet_title, "Date": fix_date(date), "Shift Time": shift_time, "Start Time": start_time, "Duration": duration})
    return results