"""
Script Name: Tico_Lib
Version: 1.0
Author: Jose Solchaga
Date: Nov 13, 2024

Description:
- 

Usage:
- 
"""

import win32com.client
from tkinter import filedialog
import tkinter as tk
import pandas as pd
import os


def open_pw(pw_file):
    """
    Open PowerWorld case and return the simauto object

    Args:
        pw_file: PowerWorld File

    Returns:
        simauto_obj: Establish PW connection
    """
    # Create PowerWorld COM instance
    simauto_obj = win32com.client.Dispatch('pwrworld.SimulatorAuto')
    
    # Open the PowerWorld case
    simauto_obj.OpenCase(pw_file)
    
    print(f"PowerWorld case '{os.path.basename(pw_file)}' opened successfully.")
    
    return simauto_obj


def close_pw(simauto_obj):
    """
    Close PowerWorld case

    Args:
        simauto_obj: PW connection

    Returns:
        None
    """
    simauto_obj.CloseCase()
    print(f"PowerWorld case closed successfully.")


def select_pw() -> str:
    """
    Ask user for single Power World case

    Args:
        None

    Returns:
        pw_file: Path to user selected PW case
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    root.attributes('-topmost', True)  # Make the root window topmost
    pw_file = filedialog.askopenfilename(title="Select PowerWorld File", filetypes=[("PowerWorld Cases", "*.pwb")])
    root.attributes('-topmost', False)  # Reset the topmost attribute
    root.destroy()  # Destroy the root window
    return pw_file


def select_aux() -> str:
    """
    Ask user for single aux file

    Args:
        None

    Returns:
        aux_file: Path to user selected aux file
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    root.attributes('-topmost', True)  # Make the root window topmost
    aux_file = filedialog.askopenfilename(title="Select Auxiliary File", filetypes=[("Auxiliary Files", "*.aux")])
    root.attributes('-topmost', False)  # Reset the topmost attribute
    root.destroy()  # Destroy the root window
    return aux_file

def select_pw_aux_multiple() -> tuple:
    """
    Ask user for multiple PW cases and aux files

    Args:
        None

    Returns:
        pw_files:  Path to user selected PW cases
        aux_files: Path to user selected aux files
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    root.attributes('-topmost', True)  # Make the root window topmost
    pw_files = filedialog.askopenfilenames(title="Select PowerWorld Files", filetypes=[("PowerWorld Cases", "*.pwb")])
    aux_files = filedialog.askopenfilenames(title="Select Aux Files", filetypes=[("Auxiliary Files", "*.aux")])
    root.attributes('-topmost', False)  # Reset the topmost attribute
    root.destroy()  # Destroy the root window
    return pw_files, aux_files


def get_file_path_info(file: str) -> tuple:
    """
    Use os.path to manipulate the path of doc path

    Args:
        String with path to file

    Returns:
        Ex: file = C:/Users/JSOLCHAGA/Downloads/23SSWG_2027_MIN_U1_Final_10092023_v88_sensitivity_V32_old.pwb
        Directory: C:/Users/JSOLCHAGA/Downloads 
        File Name: 23SSWG_2027_MIN_U1_Final_10092023_v88_sensitivity_V32_old.pwb 
        Root:      C:/Users/JSOLCHAGA/Downloads/23SSWG_2027_MIN_U1_Final_10092023_v88_sensitivity_V32_old 
        Extension: .pwb
    """
    directory = os.path.dirname(file)
    file_name = os.path.basename(file)
    root, ext = os.path.splitext(file)

    return directory, file_name, root, ext


def aux_file_applicator():
    """
    Apply auxiliary files to PowerWorld cases.
    This function processes selected PowerWorld cases and applies the corresponding auxiliary files.
    Allows the user to apply one or more aux files to multiple PW cases.

    Args:
        None

    Returns:
        None
    """
    # Select PW cases and Aux files to be applied
    pw_files, aux_files = select_pw_aux_multiple()

    # Iterate over the PW cases and apply the aux files to each case
    for case in pw_files:
        directory, file_name, root, ext = get_file_path_info(case)
        simauto_obj = open_pw(case)
        # Iterate over the aux files to be applied. Will run Power Flow twice after applying each aux file
        for aux in aux_files:
            simauto_obj.RunScriptCommand("EnterMode(EDIT);")
            print(f"Applying aux file: {os.path.basename(aux)}")
            simauto_obj.ProcessAuxFile(aux)
            simauto_obj.RunScriptCommand("SolvePowerFlow")
            simauto_obj.RunScriptCommand("SolvePowerFlow")
        # Save the log
        log = root + "_log.txt"
        simauto_obj.RunScriptCommand(f'LogSave({log}, NO)')
        print(f"Saved log: {log}")
        # Save the case with the aux files
        case_with_aux = root + "_with_aux.PWB"
        simauto_obj.SaveCase(case_with_aux, "PWB", False)
        print(f"Saved case: {case_with_aux}")
        simauto_obj.CloseCase()











if __name__ == "__main__":
    print("Testing Tico_Lib.py\n")



