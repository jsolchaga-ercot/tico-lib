import sys
import os
import win32com.client
from tkinter import filedialog
import pandas as pd

# Import Tico_Lib
sys.path.append('C://Users//JSOLCHAGA//OneDrive - ERCOT//Documents//Python//tico-lib//Tico_Lib')
import Tico_Lib as tico

############################################################################################################
# FUNCTIONS

def get_precontingency_violations(simauto_obj) -> pd.DataFrame:
    # Get all of the pre-contingency bus violations
    object_type = 'Bus'
    param_list = ['Number', 'Name', 'Vpu', 'LimitHighA', 'LimitHighB', 'ZoneName']
    filter_name = 'Pre-Contingency Violations All Regions'
    #filter_name = 'Pre-Contingency WFW Violations'

    # TODO: Figure out how to do advanced filters without the need of an advanced filter saved into the case
    '''
    zone_filter = ' or '.join([f'ZoneName = "{zone}"' for zone in weather_zone])
    voltage_filter = 'Vpu > LimitHighA'
    composite_filter = f'{zone_filter} and {voltage_filter}'
    '''

    err, output = simauto_obj.GetParametersMultipleElement(object_type, param_list, filter_name)
    df_precontingency = pd.DataFrame(output).transpose()
    if df_precontingency.empty:
        print("All pre-contingency violations solved!")
        sys.exit()
    df_precontingency.columns = param_list
    columns_to_convert = ['Vpu', 'LimitHighA', 'LimitHighB']
    df_precontingency[columns_to_convert] = df_precontingency[columns_to_convert].astype(float)

    return df_precontingency


def write_aux_cap(list_shunts):
    current_directory = os.path.dirname(__file__)
    file_name = 'PreContingency_HighVoltage_Tuning.aux'
    file_path = os.path.join(current_directory, file_name)

    with open(file_path, "a") as file:
        file.write("Shunt (BusNum,BusName,ID,MvarNom,ShuntMode)\n")
        file.write("{\n")
        for i in range(len(list_shunts)):
            file.write(f'       {list_shunts[i][0]} "{list_shunts[i][1]}"   "{list_shunts[i][2]}"     {0} "{"Fixed"}"\n')
        file.write("}\n")
        file.write("\n")


def write_aux_setpoint(list_gen):
    current_directory = os.path.dirname(__file__)
    file_name = 'PreContingency_HighVoltage_Tuning.aux'
    file_path = os.path.join(current_directory, file_name)

    with open(file_path, "a") as file:
        file.write("Gen (BusNum,BusName,ID,VoltSet)\n")
        file.write("{\n")
        for i in range(len(list_gen)):
            file.write(f'       {list_gen[i][0]} "{list_gen[i][1]}"   "{list_gen[i][2]}"  {list_gen[i][3]}\n')
        file.write("}\n")
        file.write("\n")


def cap_tuning(bus_number, bus_Vpu, bus_LimitHighA, df_shunt, simauto_obj):
    list_shunts = []
    # Loop over the reactive devices until Vpu < LimitHighA
    for i in range(len(df_shunt)):
        # Current shunt info
        shunt_bus_number = df_shunt.iloc[i]['BusNum']
        shunt_bus_name = df_shunt.iloc[i]['BusName']
        shunt_bus_id = df_shunt.iloc[i]['ID']

        # Change the value of the shunt with the highest cap bank to zero
        err = simauto_obj.ChangeParametersSingleElement('Shunt', ['BusNum', 'BusName', 'ID', 'MvarNom', 'ShuntMode'], 
                                                        [shunt_bus_number, shunt_bus_name, shunt_bus_id, 0, 'Fixed'])
        list_shunts.append([shunt_bus_number, shunt_bus_name, shunt_bus_id, 0, 'Fixed'])
        err = simauto_obj.RunScriptCommand("SolvePowerFlow(RECTNEWT)")

        # Get the voltage value of the bus with the violation we are looking to see if solved
        err, bus = simauto_obj.GetParametersSingleElement('Bus', ['Number', 'Name', 'Vpu'], [bus_number,0,0])
        bus_Vpu = float(bus[2].strip())

        if bus_Vpu < bus_LimitHighA:
            write_aux_cap(list_shunts)
            list_shunts = []
            return True

        if i > 10:
            print('Can\'t solve with cap banks. Will try set point voltages.')
            return False


def setpoint_tuning(bus_number, bus_Vpu, bus_LimitHighA, df_gen, simauto_obj):
    list_gen = []
     
    # Using Gen Set Points
    for i in range(len(df_gen)):
        gen_number = df_gen.iloc[i]['BusNum']
        gen_name = df_gen.iloc[i]['BusName']
        gen_id = df_gen.iloc[i]['ID']
        gen_setpoint = df_gen.iloc[i]['VoltSet']

        # Change set point voltage
        err = simauto_obj.ChangeParametersSingleElement('Gen', ['BusNum', 'ID', 'VoltSet'], [gen_number, gen_id, gen_setpoint*0.985])
        list_gen.append([gen_number, gen_name, gen_id, gen_setpoint*0.985])
        err = simauto_obj.RunScriptCommand("SolvePowerFlow(RECTNEWT)")

        # Get the voltage value of the bus with the violation we are looking to see if solved
        err, bus = simauto_obj.GetParametersSingleElement('Bus', ['Number', 'Name', 'Vpu'], [bus_number,0,0])
        bus_Vpu = float(bus[2].strip())

        if bus_Vpu < bus_LimitHighA:
            write_aux_setpoint(list_gen)
            list_gen = []
            return True

        if i > 10:
            print('Can\'t solve with setpoint voltages either.')
            sys.exit()
            return False    


############################################################################################################
# GET DATA

# Select and open PW case
pw_file = tico.select_pw()
#pw_file = '23SSWG_2030_SUM1_U1_Final_10092023_v96.pwb'
simauto_obj = tico.open_pw(pw_file)

weather_zone = ['WEST', 'FAR_WEST']

df_precontingency = get_precontingency_violations(simauto_obj)

############################################################################################################
# START SOLVING

print(f'There are a total of {len(df_precontingency)} pre-contingency violations in WFW.')

first = True

while not df_precontingency.empty:
    if not first:
        df_precontingency = get_precontingency_violations(simauto_obj)

    bus_number = df_precontingency.iloc[0]['Number']
    bus_name = df_precontingency.iloc[0]['Name']
    bus_Vpu = df_precontingency.iloc[0]['Vpu']
    bus_LimitHighA = df_precontingency.iloc[0]['LimitHighA']


    # Grab the first pre-contingency high voltage
    print(f'Solving {bus_name} ({bus_number}) pre-contingency high voltage of {bus_Vpu}')

    # Calculate voltage sensitivities
    err = simauto_obj.RunScriptCommand(f"CalculateVoltSense(BUS {bus_number})")

    if err == '':
        print(err)
    else:
        # Get the sensitivities of all shunts with respect to the violating bus
        err, shunt = simauto_obj.GetParametersMultipleElement('Shunt', ['BusNum', 'BusName', 'ID', 'SensdValuedQnominj', 
                                                                    'MvarNom', 'MvarNomMax', 'MvarNomMin', 'ShuntMode'], '')

        # Convert to Data Frame for easier manipulation and format data
        df_shunt = pd.DataFrame(shunt).transpose()
        df_shunt.columns = ['BusNum', 'BusName', 'ID', 'SensdValuedQnominj', 'MvarNom', 'MvarNomMax', 'MvarNomMin', 'ShuntMode']
        columns_to_convert = ['SensdValuedQnominj', 'MvarNom', 'MvarNomMax', 'MvarNomMin']
        df_shunt[columns_to_convert] = df_shunt[columns_to_convert].astype(float)

        # Sort sensitivities from highest to lowest
        df_shunt = df_shunt.sort_values(by='SensdValuedQnominj', ascending=False)

        # Get the sensitivities for Generators with respect to the violating bus
        err, gen = simauto_obj.GetParametersMultipleElement('Gen', ['BusNum', 'BusName', 'ID', 'SensdValuedVset', 
                                                                             'VoltSet'], '')
        
        # Convert to Data Frame for easier manipulation and format data
        df_gen = pd.DataFrame(gen).transpose()
        df_gen.columns = ['BusNum', 'BusName', 'ID', 'SensdValuedVset', 'VoltSet']
        columns_to_convert = ['SensdValuedVset', 'VoltSet']
        df_gen[columns_to_convert] = df_gen[columns_to_convert].astype(float)

        # Sort sensitivities from highest to lowest
        df_gen = df_gen.sort_values(by='SensdValuedVset', ascending=True)

        cap_solved = cap_tuning(bus_number, bus_Vpu, bus_LimitHighA, df_shunt, simauto_obj)

        if not cap_solved:
            setpoint_solved = setpoint_tuning(bus_number, bus_Vpu, bus_LimitHighA, df_gen, simauto_obj)



        # To avoid having to re-pull the pre-contingency violations on the first iteration
        first = False


print('T')



