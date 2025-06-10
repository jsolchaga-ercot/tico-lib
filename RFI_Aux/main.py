from tkinter import filedialog
from tkinter.filedialog import askdirectory
from datetime import datetime
import pandas as pd
import os
import math


def get_duplicates(name_list: list) -> list:
    """ Identifies the loads that are duplicated in the RFI spreadsheet

        Args:
            name_list: All the load names in the RFI spreadsheet
        
        Returns:
            duplicates: A list with the load names that are duplicated in the RFI spreadsheet
    """
    duplicates = []
    seen = set()
    
    for load in name_list:
        if load in seen:
            if load not in duplicates:  # Avoid adding the same duplicate multiple times
                duplicates.append(load)
        else:
            seen.add(load)

    return duplicates


def write_aux(aux_template: str, aux_string:str, respondent:str, year:str, load_name:str, aux_destination:str, duplicates:list, wz:str, load_type:str):
    """ Function that creates the Step 2 aux files """

    # Variable lines is a list of strings. Those strings are all the lines in the aux template + the new load information (variable aux_string)
    with open(aux_template, 'r') as file:
        lines = file.readlines()
    lines.insert(7, f"  {aux_string}")

    # The files are organized user_path/WZ/TSP/Year/Contract-OfficerLetter
    new_path = f"{aux_destination}//{wz.rstrip()}//{respondent.rstrip()}//{year}//{load_type}"
    
    # If the new path doesn't exist already, create it
    if not os.path.exists(new_path):
        os.makedirs(new_path)
    
    # Name the new file "{Data_Corrections_Updates}_TSP_LoadName_Year.aux"
    file_path = os.path.join(new_path, f"{{Data_Corrections_Updates}}_{respondent.rstrip()}_{load_name}_{year}.aux")

    # If there is already an aux file for the same load name and type (Conract or OL) but a different ID or bus, then add it to the same aux file.
    if os.path.exists(file_path) and (load_name in duplicates or 'Aggregated' in load_name):
        with open(file_path, 'r') as file:
            existing_lines = file.readlines()
        existing_lines.insert(8, f"  {aux_string}\n")
        with open(file_path, 'w') as file:
            file.writelines(existing_lines)

    # Create the aux file for Step 2
    if not os.path.exists(file_path):
            with open(file_path, 'w') as file:
                file.writelines(lines)

    return 0


def calculate_mvar(mw: float, pf: float) -> float:
    """ Calculate MVar values

        Args:
            mw: mw value listed in RFI spreadsheet
            pf: pf value listed in RFI spreadsheet
        
        Returns:
            mvar: Calculated MVar value rounded to 2 decimal places
    """
    mvar = mw * (math.tan(math.acos(pf)))
    mvar = round(mvar, 2)
    return mvar


def weather_zone(wz: str) -> str:
    """ Simplifies the weather zone information to create the weather zone folders

        Args:
            wz: The RTP weather zone in the RFI spreadsheet
        
        Returns:
            A string with the simplified weather zone
    """
    if 'west' in wz.lower():
        return 'WFW'
    elif 'coast' in wz.lower() or 'east' in wz.lower():
        return 'EC'
    elif 'north' in wz.lower() or 'ncent' in wz.lower():
        return 'NNC'
    elif 'south' in wz.lower() or 'scent' in wz.lower() or 'southern' in wz.lower():
        return 'SSC'
    else:
        return 'Unknown'
    

def main():
    # Ask user for rfi spreadsheet to process. The RFI spreadsheet should have all the important values filled like ID, County, Contract/OfficerLetter, etc.
    rfi = filedialog.askopenfilename(title="Select the RFI spreadsheet", filetypes= [("Microsoft Excel Files", ".xlsx")])
    
    # This aux file template ads the county information and designates the type of load (Contract or Officer Letter)
    aux_template = filedialog.askopenfilename(title="Select the aux template", filetypes= [("AUX File", ".aux")])
    
    # Path were the tool will place the Step 2 files
    aux_destination = askdirectory(title="Select destination folder for aux files")
    
    # Temp data frame with ALL the spreadsheet values
    temp_df = pd.read_excel(rfi)
    
    # Data Frame with only the KEY values.
    # This is were you can modify which column to use to create step 2 files.
    # The tool is currently using the ADJUSTED 3 column.
    df = temp_df[['Respondent', 'Load Name', 'Wzone', 'County',
                    'Contract/Officer Letter', 'Bus Number', 'Load ID', 'Power Factor',
                    'Summer2027 (Adj 3)', 'Winter2027-2028 (Adj 3)', 'Summer2028 (Adj 3)',
                    'Summer2030 (Adj 3)', 'Summer2031 (Adj 3)', 'Expected Start Year', 'RTP WZ']]

    # Create a list with all the load names to identify the names that are duplicated. The reason some can be duplicated is
    # becaue we splitted the loads that hade multiple buses/IDs
    name_list = df['Load Name'].tolist()
    duplicates = get_duplicates(name_list)

    # Get current date and time
    now = datetime.now().strftime("%m%d%y-%H%M%S")

    # Create folder named "AuxFiles <current date and time>" in the destination folder
    path = f"{aux_destination}//AuxFiles {now}"
    if not os.path.exists(path):
        os.makedirs(path)

    # Output message
    print(f"Creating Aux Files...")

    # Iterate over all the loads in the spreadsheet
    for index, row in df.iterrows():
        # Get the key values into variables
        pf = row['Power Factor']
        respondent = row['Respondent']
        load_name = row['Load Name']
        load_id = row['Load ID']
        county = row['County']
        bus_number = row['Bus Number']
        start_year = row['Expected Start Year']
        wz = row['Wzone']
        wz_RTP = row['RTP WZ']
        load_type = row['Contract/Officer Letter']

        # Identify the type of load: 1) Contract, 2) Officer Letter, or 3) Third Party Study (Centerpoint)
        if 'Contract' in load_type:
            load_type = 'Contract'
        elif 'Officer' in load_type:
            load_type = 'OfficerLetter'
        elif '3rd Party Study' in load_type:
            load_type = '3rdPartyStudy'
        else:
            continue

        ### Data validation START ###

        # Check for rows without the respondent (TSP) information
        if pd.isna(respondent):
            if not pd.isna(load_name):
                print(f"No respondent information for load '{load_name}'")
            else:
                print(f"No respondent or load name information.")
            continue

        # Only create Step 2 files for loads that have an energization date before 2031
        if not pd.isna(start_year):
            if int(start_year) > 2031:
                continue
        
        # Check for WZ information
        wz = weather_zone(wz_RTP)
        if wz == 'Unknown':
            print(f"Check weather zone information ('{wz}')")
            continue

        # Make sure that the load has county information
        if pd.isna(county):
            print(f"Check county information ('{county}') for {respondent}'s load '{load_name}'")
            continue

        # Check for load ID information
        if '/' in str(load_id) or pd.isna(load_id) or '?' in str(load_id):
            print(f"Check the ID value ('{load_id}') for {respondent}'s load '{load_name}' ")
            continue

        # Check for bus information
        if '/' in str(bus_number) or pd.isna(bus_number) or '?' in str(bus_number):
            print(f"Check bus number info ('{bus_number}') for {respondent}'s load '{load_name}'")
            continue

        # Special check for rows with Aggregated entries
        if 'Aggregated' in str(load_name):
            load_name = 'Aggregated Non Data-Center Large Load - 5MW to 75MW'

        ### Data validation END ###

        # Calculate the mvar values for all years
        mvar2031 = calculate_mvar(row['Summer2031 (Adj 3)'], pf)
        mvar2030 = calculate_mvar(row['Summer2030 (Adj 3)'], pf)
        mvar2028 = calculate_mvar(row['Summer2028 (Adj 3)'], pf)
        mvar2028min = calculate_mvar(row['Winter2027-2028 (Adj 3)'], pf)
        mvar2027 = calculate_mvar(row['Summer2027 (Adj 3)'], pf)

        # Create the strings that will be used in the aux template
        aux_2031 = f"{row['Bus Number']} \"{row['Load ID']}\" \"Closed\" \"YES\" {round(row['Summer2031 (Adj 3)'],2)} {mvar2031} \"{row['County']}\" \"{row['Contract/Officer Letter']}\""
        aux_2030 = f"{row['Bus Number']} \"{row['Load ID']}\" \"Closed\" \"YES\" {round(row['Summer2030 (Adj 3)'],2)} {mvar2030} \"{row['County']}\" \"{row['Contract/Officer Letter']}\""
        aux_2028 = f"{row['Bus Number']} \"{row['Load ID']}\" \"Closed\" \"YES\" {round(row['Summer2028 (Adj 3)'],2)} {mvar2028} \"{row['County']}\" \"{row['Contract/Officer Letter']}\""
        aux_2028min = f"{row['Bus Number']} \"{row['Load ID']}\" \"Closed\" \"YES\" {round(row['Winter2027-2028 (Adj 3)'],2)} {mvar2028min} \"{row['County']}\" \"{row['Contract/Officer Letter']}\""
        aux_2027 = f"{row['Bus Number']} \"{row['Load ID']}\" \"Closed\" \"YES\" {round(row['Summer2027 (Adj 3)'],2)} {mvar2027} \"{row['County']}\" \"{row['Contract/Officer Letter']}\""

        # Call a function to create the Step2 aux files 
        write_aux(aux_template, aux_2031, respondent, '2031Sum', load_name, path, duplicates, wz, load_type)
        write_aux(aux_template, aux_2030, respondent, '2030Sum', load_name, path, duplicates, wz, load_type)
        write_aux(aux_template, aux_2028, respondent, '2028Sum', load_name, path, duplicates, wz, load_type)
        write_aux(aux_template, aux_2028min, respondent, '2028Min', load_name, path, duplicates, wz, load_type)
        write_aux(aux_template, aux_2027, respondent, '2027Sum', load_name, path, duplicates, wz, load_type)

    print(f"Finished! Aux files can be found here: {path}")

main()