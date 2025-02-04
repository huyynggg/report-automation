# report-automation
A Python script which created to perform a report automation process. 

import os
import pandas as pd
import logging
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import calendar


# Configure logging
logging.basicConfig(
    filename=confidential_file_path,  
    filemode='a',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Settings for different locations
settings = {
    'Anderlecht': {'file_path': r'\\******\********\******\***\********\********'},
    'Antwerpen': {'file_path': r'\\******\********\******\***\********\********''},
    'De Pinte': {'file_path': r'\\******\********\******\***\********\********''},
    'Jumet': {'file_path': r'\\******\********\******\***\********\********''},
    'Leuven': {'file_path': r'\\******\********\******\***\********\********''},
    'Ans': {'file_path': r'\\******\********\******\***\********\********''},
    'Jabbeke': {'file_path': os.path.join(os.path.expanduser("~"), '********', '*********')}
}

# Plant codes mapping for each location
plant_codes = {
    'Anderlecht': 'BE09',
    'Antwerpen': 'BE06',
    'De Pinte': 'BE07',
    'Jumet': 'BE11',
    'Leuven': 'BE10',
    'Ans': 'BE12',
    'Jabbeke': 'BE05',
}  

# Get the current date, year, and month
now = datetime.now()
current_year = now.year
current_month = now.month
current_month_formatted = f"{current_month:02d}" 
current_year_short = current_year % 100

# Local destination file path
user_home_dir = os.path.expanduser('~')  # This gets the current user's home directory
destination_file = os.path.join(user_home_dir, r'********\***********\*******')


# Function to process each file and append data
def process_file(location_name, file_path):
    try:
        # Read the specific columns based on location
        df = pd.read_excel(file_path, sheet_name='Hours', engine='pyxlsb', usecols=[6, 13])

        # Debug: Check the raw data
        print(f"Raw data for {location_name}:")
        
        df = df.iloc[2:32, :]
        # Trimming the DataFrame to valid rows 
        num_days_in_month = calendar.monthrange(current_year, current_month)[1]
        df = df.iloc[:num_days_in_month, :]

        # Insert the custom columns
        plant_code = plant_codes.get(location_name, 'UNKNOWN')
        df.insert(0, 'Location', location_name)
        df.insert(1, 'Plant Code', plant_code)

        # Add date column for the current month
        start_date = datetime(now.year, now.month, 1)
        date_list = pd.date_range(start=start_date, periods=len(df)).strftime('%d/%m/%Y')
        df.insert(2, 'Date', date_list)

        # Debug
        print(f"Prepared DataFrame for {location_name}:")
        print(df.head())

        # Load the workbook
        wb = load_workbook(destination_file)
        sheet = wb['Hours_Test_2']

        # Find the first empty row
        first_empty_row = sheet.max_row + 1

        # Debug
        print(f"Writing data to row {first_empty_row} for {location_name}")

        # Convert the DataFrame to rows
        rows = dataframe_to_rows(df, index=False, header=False)

        # Write each row to the worksheet
        for r_idx, row in enumerate(rows, first_empty_row):
            for c_idx, value in enumerate(row, 1):
                sheet.cell(row=r_idx, column=c_idx, value=value)

        # Save the workbook
        wb.save(destination_file)

        logging.info(f"Data for {location_name} added successfully.")
        print(f"Data for {location_name} added successfully. Written to row {first_empty_row}.")
    
    except Exception as e:
        logging.error(f"An error occurred for {location_name}: {e}")
        print(f"An error occurred for {location_name}: {e}")

# Loop through each location and process the corresponding file
for location, path in settings.items():
    # Define file name based on the location
    if location == 'Leuven':
        file_name = f"LeuvenHasselt_{str(current_year)}_{str(current_month_formatted)}.xlsb"
    elif location == 'Anderlecht':
        file_name = f"Anderlecht {str(current_month_formatted)}-{str(current_year_short)}.xlsb"
    elif location == 'Antwerpen':
        file_name = f"Antwerpen {current_month_formatted}-{str(current_year)}.xlsb"
    elif location == 'De Pinte':
        file_name = f"De Pinte {current_month_formatted}-{str(current_year)}.xlsb"
    elif location == 'Jumet':
        file_name = f"Jumet {current_month_formatted}-{str(current_year)}.xlsb"
    elif location == 'Ans':
        file_name = f"Ans {current_month_formatted}.{str(current_year)}.xlsb"
    elif location == 'Jabbeke':
        file_name = f"Jabbeke {current_month_formatted}-{str(current_year)}.xlsb"

    # Add current year directory for specific locations
    if location in ['Ans', 'Anderlecht', 'Antwerpen', 'De Pinte', 'Jumet', 'Leuven']:
        file_path = os.path.join(path['file_path'], str(current_year), file_name)  
    else:
        file_path = os.path.join(path['file_path'], file_name) 

    # Check if the file exists
    if not os.path.exists(file_path):
        logging.warning(f"File not found for {location}: {file_path}")
        print(f"File not found for {location}: {file_path}")
        continue  

    # Log and process the file
    logging.info(f"Processing location: {location}")
    print(f"Processing location: {location}")

    process_file(location, file_path)
    
    

# === Merged Script Separator ===

# -*- coding: utf-8 -*-
"""
Created on Fri Sep 13 10:40:56 2024

@author: NGUYENGI
"""

import os
import pandas as pd
import logging
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import calendar

logging.basicConfig(
    filename='DSP_log.log',  # Make sure this path is valid
    filemode='a', 
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Get the current date, year, and month
now = datetime.now()
current_year = now.year
current_month = now.month
current_month_formatted = f"{current_month:02d}" 
current_year_short = current_year % 100

# Local destination file path
user_home_dir = os.path.expanduser('~')
destination_file = os.path.join(user_home_dir, r'*******\********\*******')


def clear_old_data():
    try:
        print("Function clear_old_data is being executed")
        # Log when the function starts
        logging.info(f"Started clearing old data for {current_month:02d}/{current_year}.")
        print(f"Started clearing old data for {current_month:02d}/{current_year}.")
        
        # Load the workbook and select the sheet
        wb = load_workbook(destination_file)
        sheet = wb['Hours_Test_2']  #Adjust sheet name
        print("Workbook loaded and sheet selected")

        # List to track rows to delete
        rows_to_delete = []

        # Iterate through rows starting from the second row 
        for row_idx in range(2, sheet.max_row + 1):
            date_value = sheet.cell(row=row_idx, column=3).value 
            # Print the date being processed
            print(f"Processing row {row_idx}, date value: {date_value}")

            # Ensure the value in the date column is a valid date
            if isinstance(date_value, str): 
                try:
                    # Convert the string date to a datetime object, assuming the format is 'dd/mm/yyyy'
                    date_obj = datetime.strptime(date_value, '%d/%m/%Y')

                    # Check if the date matches the current month and year
                    if date_obj.month == current_month and date_obj.year == current_year:
                  
                        rows_to_delete.append(row_idx)

                except ValueError:
                    
                    print(f"Invalid date format in row {row_idx}: {date_value}")
                    logging.warning(f"Invalid date format found in row {row_idx}: {date_value}")
                    continue

        # Print the rows marked for deletion
        print(f"Rows to delete: {rows_to_delete}")
        logging.info(f"Rows to delete: {rows_to_delete}")

        # Delete rows in reverse order to avoid shifting row indexes
        for row_idx in reversed(rows_to_delete):
            sheet.delete_rows(row_idx)

        # Save the workbook after deletion
        wb.save(destination_file)

        # Log the result of the deletion
        if rows_to_delete:
            logging.info(f"Data for {current_month:02d}/{current_year} cleared successfully.")
            print(f"Data for {current_month:02d}/{current_year} cleared successfully.")
        else:
            logging.info(f"No data found for {current_month:02d}/{current_year} to delete.")
            print(f"No data found for {current_month:02d}/{current_year} to delete.")

    except Exception as e:
        logging.error(f"Error clearing old data for {current_month:02d}/{current_year}: {e}")
        print(f"Error clearing old data for {current_month:02d}/{current_year}: {e}")

if __name__ == "__main__":
    clear_old_data()
