import pandas as pd
import math

excel_file = r"C:\Users\aganea\Documents\Project_Elevate_BPML (3).xlsx"
sheet_name = "SteriMax User Mapping"
df = pd.read_excel(excel_file, sheet_name=sheet_name)

for col in df:
    df[col]=df[col].astype(str).apply(lambda x: x.replace('x',df[col].name))

# print(df)

# Define an empty dictionary to store the data
data_dict = {}

# Iterate over the rows in the DataFrame
for index, row in df.iterrows():
    # Create a dictionary entry for each row
    key = row[df.columns[1]]
    row_dict = {}
    for column in df.columns[2:]:
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
    # print(f"Key: {key}")
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

for dic in employee_list: # Iterates through the master list of dictionaries and sees if someone has a specific role, then adds them to one of the three above lists
    for key, employee_value in dic.items():
        if 'Z_FLP_USER' in employee_value:
            advanced_user_list.append(key)

print(employee_list)
print(advanced_user_list)





