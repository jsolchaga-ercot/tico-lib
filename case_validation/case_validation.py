import sys
import os
import win32com.client
from tkinter import filedialog
import pandas as pd

# Import Tico_Lib
sys.path.append('C://Users//JSOLCHAGA//OneDrive - ERCOT//Documents//Python//tico-lib//Tico_Lib')
import Tico_Lib as tico

pw_file = tico.select_pw()
simauto_obj = tico.open_pw(pw_file)

############################################################################################################
# P_Gen > P_Max and P_Min < P_Min

# Parameters
object_list = 'Gen'
param_list = ['BusNum', 'BusName', 'ID', 'MW', 'MWMax', 'MWMin']
filter_name = ''

# Get table information
err, output = simauto_obj.GetParametersMultipleElement(object_list, param_list, filter_name)

# Convert to Data Frame
df_gen = pd.DataFrame(output).transpose()
df_gen.columns = param_list


############################################################################################################