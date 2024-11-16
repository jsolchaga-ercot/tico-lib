from tkinter import filedialog
import win32com.client
from pyrtp.pwd import pwd
import pandas as pd



def main():
    # User Inputs PW Case and Aux File
    pw_file = filedialog.askopenfilename(title="Select your PowerWorld case", filetypes= [("PowerWorld files", "*.pwb")])
    aux_file = filedialog.askopenfilename(title="Select aux file to create backout file", filetypes=[("Aux Files", "*.aux")])
    path = aux_file.split('.')[0] + '_BACKOUT.aux'
    # Create PowerWorld COM instance
    simauto_obj = win32com.client.Dispatch('pwrworld.SimulatorAuto')
    simauto_obj.OpenCase(pw_file)

    # Apply aux file to case
    simauto_obj.RunScriptCommand("EnterMode(EDIT);")
    simauto_obj.ProcessAuxFIle(aux_file)
    simauto_obj.RunScriptCommand('DiffCaseSetAsBase')
    simauto_obj.CloseCase()

    # Run Difference Case
    simauto_obj.OpenCase(pw_file)
    simauto_obj.RunScriptCommand(f"DiffFlowWriteCompleteModel(\"{path}\", YES, YES, YES, YES, PRIMARY, \"Network Model\",,,,\"NO\");")
    


main()