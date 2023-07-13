import math
import os
import io
import sys
import glob
import time
from datetime import datetime
import tempfile
import shutil

#### Importing Local Libraries
root_dir = os.path.dirname(os.path.abspath(__file__))
root_dir = root_dir.replace("\\", "/")
libraries_dir = root_dir + "/libs"
sys.path.append(libraries_dir) # Used to import local pip libraries
import pandas as pd
import openpyxl
import xlsxwriter
import sharepy

site_url = "https://sterimax.sharepoint.com"

s = sharepy.connect(site_url)

file_url = "https://sterimax.sharepoint.com/:x:/r/sites/SMIProjectElevate/Shared%20Documents/1.3%20Realize%20Build/BPML/Project_Elevate_BPML.xlsx"
download_path = "./spreadsheets/input/Project_Elevate_BPML.xlsx"

r = s.getfile(file_url, filename=download_path)
 
df = pd.read_excel(download_path)


def find_common_element(list1, list2):  # Function that checks if an employee has only 1 role in a certain category
    common_elements = set(list1) & set(list2)  # Find common elements using set intersection
    return len(common_elements) == 1 # Returns TRUE if and only if the lists have one element in common

def count_common_elements(list1, list2): # Function that checks how many roles a certain user has of a category
    common_elements = set(list1) & set(list2)  # Find common elements using set intersection
    return len(common_elements)

def get_library_name(): # Used to easily obtain the path of the libraries
    root_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = root_dir.replace("\\", "/")
    libraries_directory = root_dir + "/libs"
    return libraries_directory

def get_input_spreadsheets_name(): # Used to easily obtain the path of the spreadsheets
    root_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = root_dir.replace("\\", "/")
    spreadsheet_directory = root_dir + "/spreadsheets/input"
    return spreadsheet_directory

def get_output_spreadsheets_name(): # Used to easily obtain the path of the spreadsheets
    root_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = root_dir.replace("\\", "/")
    spreadsheet_directory = root_dir + "/spreadsheets/output"
    return spreadsheet_directory


##### Contains the logic to pick an Excel file and a valid sheet.

input_spreadsheet_directory = get_input_spreadsheets_name()

excel_files = glob.glob(os.path.join(input_spreadsheet_directory, '*.xlsx')) # Looks at all the .xlsx files in the directory

if len(excel_files) == 0: # Closes the program is the user forgot to put in an input file.
    print("No Excel files found, please re-run program with your input Excel file in the correct folder.")
    exit()

print("Found Excel file(s):")
for i, file in enumerate(excel_files):
    print(f"{i+1}. {file}") # Prints out all found files.

while True: # Error handling for choosing a correct Excel file
    choice = input("Enter the number of the Excel file you want to choose (or 'q' to quit): ")
    if choice.lower() == 'q': # Allows the user to quit program
        exit()
    try:
        choice = int(choice)
        if 1 <= choice <= len(excel_files):
            excel_file = excel_files[choice - 1]
            print(f"You selected: {excel_file}")
            break
        else:
            print("Invalid choice. Please enter a valid number.")
    except ValueError:
        print("Invalid choice. Please enter a valid number.")

######


#### Code that allows a user to select a specific sheet where the BPML sheet is stored.
excel_file_name = pd.ExcelFile(excel_file)
sheet_names = excel_file_name.sheet_names

print("Available sheet(s):")
for index, name in enumerate(sheet_names):
    print(f"{index + 1}. {name}")

while True:
    try:
        sheet_index = int(input("Enter the index of the sheet you want to use that contains your users and the roles that they have. 1 is the first sheet, 2 is the second sheet, etc.: ")) - 1
        if sheet_index < 0 or sheet_index >= len(sheet_names):
            raise ValueError("Invalid sheet index. Please Try Again")
        
        break
    except ValueError:
        print("Invalid input. Please enter a valid integer.")

########


### Turn the selected sheet into a dictionary of dictionaries (each internal dictionary is a row)
selected_sheet = sheet_names[sheet_index]
df = pd.read_excel(excel_file, sheet_name=selected_sheet)

for col in df: # Replace any x in the Excel file with the column's name
    df[col]=df[col].astype(str).apply(lambda x: x.replace('x',df[col].name))

# Define an empty dictionary to store the data
data_dict = {}

# Iterate over the rows in the Excel File
for index, row in df.iterrows():
    # Create a dictionary entry for each row
    key = row[df.columns[1]] # Keys are the names of the employees, read from the second column in Excel
    row_dict = {}
    for column in df.columns[2:]: # Checks for values in the 3rd column onwards
        value = row[column]
        row_dict[column] = value
    # Add the row dictionary to the data dictionary
   
    data_dict[key] = row_dict
#####


# Extract each internal dictionary with associated key
extracted_dicts = [(key, value) for key, value in data_dict.items()]
employee_list = []

# Create a list of dictionaries with each dictionary containing the employee and their roles
for key, internal_dict in extracted_dicts:
    value_list = [(k, v) for k, v in internal_dict.items()] # Turn dictionary into a list of tuples
    cleaned_value_list = list(filter(lambda x: x[1] != 'nan', value_list)) # Clean the list of any tuples that have 'nan'
    dissolved_value_list = [item for t in cleaned_value_list for item in t] # Dissolve the tuples into a list
    
    list_before_dictionary_entry = []
    

    for item in dissolved_value_list: # Remove all duplicates from the list
        if item not in list_before_dictionary_entry:
            list_before_dictionary_entry.append(item)
    
    employee_dict = {key: list_before_dictionary_entry} # Create a separate dictionary for each employee and their roles
    employee_list.append(employee_dict) # Add each employee to a master list of dicionaries

####### Initializing all lists used to store data
advanced_user_list = []
core_user_list = []
self_service_list = []

advanced_user_constant_list = []
core_user_constant_list = []
self_service_constant_list = []

only_one_advanced_user_list = []
only_one_core_user_list = []
only_one_self_service_user_list = []

employee_list_with_counts = []

total_FUE_export_list = []

PO_CREATOR = ["ZS_MM_PO_CREATOR"]
po_creator_users_list = []
#######


#### Code that allows the user to select where the Role Mapping sheet is.
excel_file_name = pd.ExcelFile(excel_file)
sheet_names = excel_file_name.sheet_names

print("Available sheet(s):")
for index, name in enumerate(sheet_names):
    print(f"{index + 1}. {name}")

while True:
    try:
        sheet_index = int(input("Enter the index of the sheet you want to use that contains data that maps roles to their category. 1 is the first sheet, 2 is the second sheet, etc.: ")) - 1
        if sheet_index < 0 or sheet_index >= len(sheet_names):
            raise ValueError("Invalid sheet index. Please Try Again")
        
        break
    except ValueError:
        print("Invalid input. Please enter a valid integer.")

selected_sheet = sheet_names[sheet_index]

df2 = pd.read_excel(excel_file, sheet_name=selected_sheet)

########

data_dict_roles = {}

# Iterate over the rows in the sheet that contains the role mapping
for index, row in df2.iterrows():
    # Create a dictionary entry for each row
    key = row[df2.columns[0]]
    row_dict = {}
    for column in df2.columns[1:]:
        value = row[column]
       
    # Add the row dictionary to the data dictionary
   
    data_dict_roles[key] = value


for key, employee_value in data_dict_roles.items(): # Create three lists of roles that will be used to determine which employees are in which role
    if employee_value == 'Advanced':
        advanced_user_constant_list.append(key)
    elif employee_value == 'Core Use':
        core_user_constant_list.append(key)
    else:
        self_service_constant_list.append(key)

for dic in employee_list: # Iterates through the master list of dictionaries and sees if someone has a specific role, then adds them to one of the three above lists
    for key, employee_value in dic.items():
        if any(x in advanced_user_constant_list for x in employee_value) == True:
            advanced_user_list.append(key)
        elif any(x in core_user_constant_list for x in employee_value) == True:
            core_user_list.append(key)
        else:
            self_service_list.append(key)

for dic in employee_list: # Iterates through the master list of dictionaries and sees if someone has one and only one role in a certain category and adds them to another, separate list
    for key, employee_value in dic.items():
        advanced_user_result = find_common_element(advanced_user_constant_list, employee_value)
        core_user_result = find_common_element(core_user_constant_list, employee_value)
        self_service_user_result = find_common_element(self_service_constant_list, employee_value)
        if advanced_user_result:
            only_one_advanced_user_list.append(key)
        if core_user_result:
            only_one_core_user_list.append(key)
        if self_service_user_result:
            only_one_self_service_user_list.append(key)

for dic in employee_list:
    for key, employee_value in dic.items():
        po_creator_user_result = find_common_element(PO_CREATOR, employee_value)
        if po_creator_user_result:
            po_creator_users_list.append(key)


for dic in employee_list: # Iterates through the master list of dictionaries and sees how many roles of each type someone has
    for key, employee_value in dic.items():
        advanced_user_count = count_common_elements(advanced_user_constant_list, employee_value)
        core_user_count = count_common_elements(core_user_constant_list, employee_value)
        self_service_user_count = count_common_elements(self_service_constant_list, employee_value)

        employee_count_dict = {}

        # Create a dictionary with an employee and how many roles they have of each type

        employee_count_dict["Name"] = key
        employee_count_dict["Advanced User Roles"] = advanced_user_count
        employee_count_dict["Core User Roles"] = core_user_count
        employee_count_dict["Self Service Roles"] = self_service_user_count

        employee_list_with_counts.append(employee_count_dict)


### Calculate total FUE
total_FUE = (len(advanced_user_list) * 1) + (len(core_user_list) * 0.2) + (len(self_service_constant_list) * 0.0333) # Calculate Total FUE based on doc rules
total_FUE_export_list.append(total_FUE)


## Exporting Data to Another Excel Sheet

output_spreadsheet_directory = get_output_spreadsheets_name()
current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file_name = f"user_output_{current_time}.xlsx"

df = pd.DataFrame()

excel_writer = pd.ExcelWriter(os.path.join(output_spreadsheet_directory, output_file_name), engine="xlsxwriter")
df.to_excel(excel_writer, index=False, sheet_name='User Output')
excel_writer.close()


excel_file = os.path.join(output_spreadsheet_directory, output_file_name)

##


# Write Counts for Each Employee
df = pd.DataFrame(employee_list_with_counts)
start_row = 0
start_col = 0

writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df.to_excel(writer, sheet_name='User Output', index=False, startrow=start_row, startcol=start_col)
worksheet = writer.sheets['User Output']
for idx, col in enumerate(df):  # loop through all columns to auto-adjust widths
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.close()

# Write The Lists to the Next Columns


existing_data = pd.read_excel(excel_file)
start_col = existing_data.shape[1]
df1 = pd.DataFrame({"Advanced Users": advanced_user_list})
df1 = pd.concat([existing_data, df1], axis=1)
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df1.to_excel(writer, sheet_name='User Output', index=False)
worksheet = writer.sheets['User Output']
for idx, col in enumerate(df1):  # loop through all columns to auto-adjust widths
        series = df1[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.close()

existing_data = pd.read_excel(excel_file)
start_col = existing_data.shape[1]
df2 = pd.DataFrame({"Core Users": core_user_list})
df2 = pd.concat([existing_data, df2], axis=1)
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df2.to_excel(writer, sheet_name='User Output', index=False)
worksheet = writer.sheets['User Output']
for idx, col in enumerate(df2):  # loop through all columns to auto-adjust widths
        series = df2[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.close()

existing_data = pd.read_excel(excel_file)
start_col = existing_data.shape[1]
df3 = pd.DataFrame({"Self-Service Users": self_service_list})
df3 = pd.concat([existing_data, df3], axis=1)
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df3.to_excel(writer, sheet_name='User Output', index=False)
worksheet = writer.sheets['User Output']
for idx, col in enumerate(df3):  # loop through all columns to auto-adjust widths
        series = df3[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.close()

existing_data = pd.read_excel(excel_file)
start_col = existing_data.shape[1]
df4 = pd.DataFrame({"Advanced Users with Only One Role": only_one_advanced_user_list})
df4 = pd.concat([existing_data, df4], axis=1)
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df4.to_excel(writer, sheet_name='User Output', index=False)
worksheet = writer.sheets['User Output']
for idx, col in enumerate(df4):  # loop through all columns to auto-adjust widths
        series = df4[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.close()


existing_data = pd.read_excel(excel_file)
start_col = existing_data.shape[1]
df5 = pd.DataFrame({"Core Users with Only One Role": only_one_core_user_list})
df5 = pd.concat([existing_data, df5], axis=1)
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df5.to_excel(writer, sheet_name='User Output', index=False)
worksheet = writer.sheets['User Output']
for idx, col in enumerate(df5):  # loop through all columns to auto-adjust widths
        series = df5[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.close()

existing_data = pd.read_excel(excel_file)
start_col = existing_data.shape[1]
df6 = pd.DataFrame({"Self-Service Users with Only One Role": only_one_self_service_user_list})
df6 = pd.concat([existing_data, df6], axis=1)
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df6.to_excel(writer, sheet_name='User Output', index=False)
worksheet = writer.sheets['User Output']
for idx, col in enumerate(df6):  # loop through all columns to auto-adjust widths
        series = df6[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.close()

existing_data = pd.read_excel(excel_file)
start_col = existing_data.shape[1]
df7 = pd.DataFrame({"People that have ZS_MM_PO_CREATOR": po_creator_users_list})
df7 = pd.concat([existing_data, df7], axis=1)
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df7.to_excel(writer, sheet_name='User Output', index=False)
worksheet = writer.sheets['User Output']
for idx, col in enumerate(df7):  # loop through all columns to auto-adjust widths
        series = df7[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.close()


existing_data = pd.read_excel(excel_file)
start_col = existing_data.shape[1]
df8 = pd.DataFrame({"Total FUE(s) Used by the Organization": total_FUE_export_list})
df8 = pd.concat([existing_data, df8], axis=1)
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df8.to_excel(writer, sheet_name='User Output', index=False)
worksheet = writer.sheets['User Output']
for idx, col in enumerate(df8):  # loop through all columns to auto-adjust widths
        series = df8[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.close()


### Generate Report of All Users and Their Roles

employee_list_for_first_report = [(k, v) for d in employee_list for k, v1 in d.items() for v in v1]

df9 = pd.DataFrame(employee_list_for_first_report, columns=['Staff Member', 'Assigned Roles'])

df9['Staff Member'] = df9['Staff Member'].mask(df9["Staff Member"].duplicated(), '')

writer = pd.ExcelWriter(excel_file, engine='openpyxl', mode='a')

df9.to_excel(writer, sheet_name='Report 1', startrow=1, index=False)

# Get the workbook object
workbook  = writer.book

# Get the worksheet by name
worksheet = workbook['Report 1']

# Write the current date in the first cell
worksheet['A1'] = "Date: " + datetime.now().strftime('%Y-%m-%d')
dims = {}
for row in worksheet.rows:
    for cell in row:
        if cell.value:
            dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))    
for col, value in dims.items():
    worksheet.column_dimensions[col].width = value

writer.close()



print(f"Data exported to '{excel_file}' successfully.")








