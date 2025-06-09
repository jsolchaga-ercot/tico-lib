from tkinter import filedialog
from tkinter.filedialog import askdirectory
from datetime import datetime
import pandas as pd
import os
import sys
import math


def get_duplicates(name_list: list) -> list:
    duplicates = []
    seen = set()

    for item in name_list:
        if item in seen:
            if item not in duplicates:  # Avoid adding the same duplicate multiple times
                duplicates.append(item)
        else:
            seen.add(item)

    return duplicates


def write_aux(aux_template: str, aux_string:str, respondent:str, year:str, load_name:str, aux_destination:str, duplicates:list):
    with open(aux_template, 'r') as file:
        lines = file.readlines()
    lines.insert(7, f"  {aux_string}")

    new_path = f"{aux_destination}//{respondent.rstrip()}//{year}"
    if not os.path.exists(new_path):
        os.makedirs(new_path)
    file_path = os.path.join(new_path, f"{{Data_Corrections_Update}}_{respondent.rstrip()}_{load_name}_{year}.aux")
    
    if os.path.exists(file_path) and (load_name in duplicates or 'Aggregated' in load_name):
        with open(file_path, 'r') as file:
            existing_lines = file.readlines()
        existing_lines.insert(8, f"  {aux_string}\n")
        with open(file_path, 'w') as file:
            file.writelines(existing_lines)

    if not os.path.exists(file_path):
        with open(file_path, 'w') as file:
            file.writelines(lines)

    #print(f"{respondent}_{load_name}_{year}.aux")

    return 0


def calculate_mvar(mw: float, pf: float) -> float:
    mvar = mw * (math.tan(math.acos(pf)))
    mvar = round(mvar, 2)
    return mvar


rfi = filedialog.askopenfilename(title="Select the RFI spreadsheet", filetypes= [("Microsoft Excel Files", ".xlsx")])
aux_template = filedialog.askopenfilename(title="Select the aux template", filetypes= [("AUX File", ".aux")])
aux_destination = askdirectory(title="Select destination folder for aux files")
temp_df = pd.read_excel(rfi)
df = temp_df[['Respondent', 'Load Name', 'Wzone', 'County',
                'Contract/Officer Letter', 'Bus Number', 'Load ID', 'Power Factor',
                'Summer2027 (Adj 2)', 'Winter2027-2028 (Adj 2)', 'Summer2028 (Adj 2)',
                'Summer2030 (Adj 2)', 'Summer2031 (Adj 2)', 'Expected Start Year']]

name_list = df['Load Name'].tolist()
duplicates = get_duplicates(name_list)

now = datetime.now()
now = now.strftime("%m%d%y-%H%M%S")
path = f"{aux_destination}//AuxFiles {now}"
if not os.path.exists(path):
    os.makedirs(path)

print("Creating Aux Files...")

for index, row in df.iterrows():
    pf = row['Power Factor']
    respondent = row['Respondent']
    load_name = row['Load Name']
    load_id = row['Load ID']
    county = row['County']
    bus_number = row['Bus Number']
    start_year = row['Expected Start Year']

    if not pd.isna(start_year):
        if int(start_year) > 2031:
            continue

    if pd.isna(county):
        print(f"Check county information ('{county}') for {respondent}'s load '{load_name}'")
        continue

    if '/' in str(load_id) or pd.isna(load_id) or '?' in str(load_id):
        print(f"Check the ID value ('{load_id}') for {respondent}'s load '{load_name}' ")
        continue

    if '/' in str(bus_number) or pd.isna(bus_number) or '?' in str(bus_number):
        print(f"Check bus number info ('{bus_number}') for {respondent}'s load '{load_name}'")
        continue

    if 'Aggregated' in str(load_name):
        load_name = 'Aggregated Non Data-Center Large Load - 5MW to 75MW'

    mvar2031 = calculate_mvar(row['Summer2031 (Adj 2)'], pf)
    mvar2030 = calculate_mvar(row['Summer2030 (Adj 2)'], pf)
    mvar2028 = calculate_mvar(row['Summer2028 (Adj 2)'], pf)
    mvar2028min = calculate_mvar(row['Winter2027-2028 (Adj 2)'], pf)
    mvar2027 = calculate_mvar(row['Summer2027 (Adj 2)'], pf)

    aux_2031 = f"{row['Bus Number']} \"{row['Load ID']}\" \"Closed\" \"YES\" {round(row['Summer2031 (Adj 2)'],2)} {mvar2031} \"{row['County']}\" \"{row['Contract/Officer Letter']}\""
    aux_2030 = f"{row['Bus Number']} \"{row['Load ID']}\" \"Closed\" \"YES\" {round(row['Summer2030 (Adj 2)'],2)} {mvar2030} \"{row['County']}\" \"{row['Contract/Officer Letter']}\""
    aux_2028 = f"{row['Bus Number']} \"{row['Load ID']}\" \"Closed\" \"YES\" {round(row['Summer2028 (Adj 2)'],2)} {mvar2028} \"{row['County']}\" \"{row['Contract/Officer Letter']}\""
    aux_2028min = f"{row['Bus Number']} \"{row['Load ID']}\" \"Closed\" \"YES\" {round(row['Winter2027-2028 (Adj 2)'],2)} {mvar2028min} \"{row['County']}\" \"{row['Contract/Officer Letter']}\""
    aux_2027 = f"{row['Bus Number']} \"{row['Load ID']}\" \"Closed\" \"YES\" {round(row['Summer2027 (Adj 2)'],2)} {mvar2027} \"{row['County']}\" \"{row['Contract/Officer Letter']}\""

    write_aux(aux_template, aux_2031, respondent, '2031Sum', load_name, path, duplicates)
    write_aux(aux_template, aux_2030, respondent, '2030Sum', load_name, path, duplicates)
    write_aux(aux_template, aux_2028, respondent, '2028Sum', load_name, path, duplicates)
    write_aux(aux_template, aux_2028min, respondent, '2028Min', load_name, path, duplicates)
    write_aux(aux_template, aux_2027, respondent, '2027Sum', load_name, path, duplicates)


print("Finished!")