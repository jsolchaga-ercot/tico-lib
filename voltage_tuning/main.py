import sys
import os
import win32com.client
from tkinter import filedialog
import pandas as pd

# Import Tico_Lib
sys.path.append('C://Users//JSOLCHAGA//OneDrive - ERCOT//Documents//Python//tico-lib//Tico_Lib')
import Tico_Lib as tico

# Select and open PW case
pw_file = tico.select_pw()
#pw_file = '23SSWG_2030_SUM1_U1_Final_10092023_v96.pwb'
simauto_obj = tico.open_pw(pw_file)

# Get results from 'Bus' Table
object_type = 'Bus'
param_list = ['Number', 'Name', 'NomkV', 'CTGViol','Vpu', 'CTGMaxVolt', 'CTGMinVolt', 'LimitHighA', 
              'LimitHighB', 'LimitLowA', 'LimitLowB', 'ViolatedNormal', 'CTGMaxVoltName', 'CTGMinVoltName', 
              'ZoneName', 'ZoneNumber']
filter_name = 'CTGViol > 0'
err, output = simauto_obj.GetParametersMultipleElement(object_type, param_list, filter_name)

# Convert to Data Frame
df_bus = pd.DataFrame(output).transpose()
df_bus.columns = param_list

# New Data Frame for Low Voltages
df_lowbus = pd.DataFrame(columns=param_list)
df_lowbus = pd.concat([df_lowbus, df_bus[df_bus['CTGMinVolt'].notna()]], ignore_index=True)

# New Data Frame for High Voltages
df_highbus = pd.DataFrame(columns=param_list)
df_highbus = pd.concat([df_highbus, df_bus[df_bus['CTGMaxVolt'].notna()]], ignore_index=True)

columns_to_convert = ['NomkV']
df_highbus[columns_to_convert] = df_highbus[columns_to_convert].astype(float)
df_highbus = df_highbus[df_highbus['NomkV'] > 100]

'''
# Get results from 'ViolationCTG' Table
object_type = 'ViolationCTG'
param_list = ['CTG_Name', 'LV_Element', 'LV_Type', 'LV_Value', 'LV_Limit', 'LV_Percent']
filter_name = 'LV_Type = "Bus High Volts"'
err, output = simauto_obj.GetParametersMultipleElement(object_type, param_list, filter_name)
# Convert to Data Frame
df_highVoltage = pd.DataFrame(output).transpose()
df_highVoltage.columns = param_list
'''

contingencies = []
for i in range(len(df_highbus)):
    ctg = df_highbus.iloc[i]['CTGMaxVoltName']
    if ctg not in contingencies:
        contingencies.append(ctg)
        simauto_obj.RunScriptCommand("CTGSetAsReference")
        simauto_obj.RunScriptCommand(f'CTGSolve("{ctg}")')
        err, bustable = simauto_obj.GetParametersMultipleElement('Bus', ['Number', 'Name', 'Vpu', 'LimitHighB', 'LimitLowB', 'ZoneName'], '')
        

    print('T')

print('T')







'''


print('T')


'''


