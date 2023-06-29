import pandas as pd
import math
import openpyxl
import xlsxwriter

def find_common_element(list1, list2):  # Function that checks if an employee has only 1 role in a certain category
    common_elements = set(list1) & set(list2)  # Find common elements using set intersection
    return len(common_elements) == 1 # Returns TRUE if and only if the lists have one element in common

def count_common_elements(list1, list2):
    common_elements = set(list1) & set(list2)  # Find common elements using set intersection
    return len(common_elements)


excel_file = r"C:\Users\aganea\Documents\Project_Elevate_BPML (3).xlsx"
sheet_name = "SteriMax User Mapping"
df = pd.read_excel(excel_file, sheet_name=sheet_name)

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


sheet_name = "Role Mapping"
df2 = pd.read_excel(excel_file, sheet_name=sheet_name)

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

for dic in employee_list: # Iterates through the master list of dictionaries and sees how many roles of each type someone has
    for key, employee_value in dic.items():
        advanced_user_count = count_common_elements(advanced_user_constant_list, employee_value)
        core_user_count = count_common_elements(core_user_constant_list, employee_value)
        self_service_user_count = count_common_elements(self_service_constant_list, employee_value)

        employee_count_dict = {}

        employee_count_dict["Name"] = key
        employee_count_dict["Advanced User Roles"] = advanced_user_count
        employee_count_dict["Core User Roles"] = core_user_count
        employee_count_dict["Self Service Roles"] = self_service_user_count

        employee_list_with_counts.append(employee_count_dict)



total_FUE = (len(advanced_user_list) * 1) + (len(core_user_list) * 0.2) + (len(self_service_constant_list) * 0.0333) # Calculate Total FUE based on doc rules
print("The total FUE used by the organization is: " + str(total_FUE))
# print(only_one_advanced_user_list)
# print(only_one_core_user_list)
# print(only_one_self_service_user_list)
print(employee_list_with_counts)


## Exporting Data to Another Excel Sheet

excel_file = r"C:\Users\aganea\Documents\output.xlsx"

# Wipe current contents to prepare for overwrite
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active
sheet.delete_rows(1, sheet.max_row)
workbook.save(excel_file)

# Write Counts for Each Employee
df = pd.DataFrame(employee_list_with_counts)
start_row = 0
start_col = 0

writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=start_row, startcol=start_col)
worksheet = writer.sheets['Sheet1']
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
df1.to_excel(writer, sheet_name='Sheet1', index=False)
worksheet = writer.sheets['Sheet1']
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
df2.to_excel(writer, sheet_name='Sheet1', index=False)
worksheet = writer.sheets['Sheet1']
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
df3.to_excel(writer, sheet_name='Sheet1', index=False)
worksheet = writer.sheets['Sheet1']
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
df4.to_excel(writer, sheet_name='Sheet1', index=False)
worksheet = writer.sheets['Sheet1']
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
df5.to_excel(writer, sheet_name='Sheet1', index=False)
worksheet = writer.sheets['Sheet1']
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
df6.to_excel(writer, sheet_name='Sheet1', index=False)
worksheet = writer.sheets['Sheet1']
for idx, col in enumerate(df6):  # loop through all columns to auto-adjust widths
        series = df6[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.close()


print(f"Data exported to '{excel_file}' successfully.")








