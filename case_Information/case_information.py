import sys
import os
import win32com.client
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt

# Import Tico_Lib
sys.path.append('C://Users//JSOLCHAGA//OneDrive - ERCOT//Documents//Python//tico-lib//Tico_Lib')
import Tico_Lib as tico

# Select and open PW case
pw_file = tico.select_pw()
simauto_obj = tico.open_pw(pw_file)

############################################################################################################
# GEN TABLE

# Parameters
object_type = 'Gen'
param_list = ['BusNum', 'ID', 'Status', 'AGC','MW', 'Mvar', 'ZoneName']
filter_name = ''

# Get case Load information
err, output = simauto_obj.GetParametersMultipleElement(object_type, param_list, filter_name)

# Convert to Data Frame
df_gen = pd.DataFrame(output).transpose()
df_gen.columns = param_list

# Convert str to int and round up
columns_to_convert = ['MW', 'Mvar']
df_gen[columns_to_convert] = df_gen[columns_to_convert].astype(float)
df_gen = df_gen.round(2)


############################################################################################################
# LOAD TABLE
 
# Parameters
object_type = 'Load'
param_list = ['BusNum', 'ID', 'Status', 'AGC','MW', 'Mvar', 'ZoneName', 'AreaName']
filter_name = ''

# Get case Load information
err, output = simauto_obj.GetParametersMultipleElement(object_type, param_list, filter_name)

# Convert to Data Frame
df_load = pd.DataFrame(output).transpose()
df_load.columns = param_list

# Convert str to int and round up
columns_to_convert = ['MW', 'Mvar']
df_load[columns_to_convert] = df_load[columns_to_convert].astype(float)
df_load = df_load.round(2)

############################################################################################################
# ZONE TABLE

# Parameters
object_type = 'Zone'
param_list = ['Number', 'Name', 'LoadMW', 'LoadMvar', 'GenMW', 'GenMvar']
filter_name = ''

# Ger case Zone Information
err, output = simauto_obj.GetParametersMultipleElement(object_type, param_list, filter_name)

# Convert to DataFrame
df_zone = pd.DataFrame(output).transpose()
df_zone.columns = param_list

# Convert str to int and round up
columns_to_convert = ['LoadMW', 'LoadMvar', 'GenMW', 'GenMvar']
df_zone[columns_to_convert] = df_zone[columns_to_convert].astype(float)
df_zone = df_zone.round(2)

# Plot
df_zone.plot(kind='bar', x='Name', y='LoadMW', title='Load MW by Zone')
plt.show()

############################################################################################################



############################################################################################################

print('T')
