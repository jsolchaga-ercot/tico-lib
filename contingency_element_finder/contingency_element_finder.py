from tkinter import filedialog
import win32com.client
import pandas as pd



def ctg(df: pd.DataFrame, element_type: int, element_buses: str, element_id: str) -> list:
    
    contingencies = []

    if element_type == 1:
        for i in range(len(df)):
            object_action = df.iloc[i,1].split(" ")
            if 'BRANCH' in object_action and element_buses[0] in object_action and element_buses[1] in object_action and element_id in object_action:
                contingencies.append(df.iloc[i,0])
            else:  
                continue
    elif element_type == 2:
        for i in range(len(df)):
            object_action = df.iloc[i,1].split(" ")
            if '3WXFORMER' in object_action and element_buses[0] in object_action and element_buses[1] in object_action and element_buses[2] in object_action and element_id in object_action:
                contingencies.append(df.iloc[i,0])
            else:  
                continue
    elif element_type == 3:
        for i in range(len(df)):
            object_action = df.iloc[i,1].split(" ")
            if 'SHUNT' in object_action and element_buses[0] in object_action and element_id in object_action:
                contingencies.append(df.iloc[i,0])
            else:  
                continue
    elif element_type == 4:
        for i in range(len(df)):
            object_action = df.iloc[i,1].split(" ")
            if 'BUS' in object_action and element_buses[0] in object_action and element_id in object_action:
                contingencies.append(df.iloc[i,0])
            else:  
                continue
    elif element_type == 5:
        for i in range(len(df)):
            object_action = df.iloc[i,1].split(" ")
            if 'GEN' in object_action and element_buses[0] in object_action and element_id in object_action:
                contingencies.append(df.iloc[i,0])
            else:  
                continue
    else:
        print("Entered number outside 1-5")
        exit

    return contingencies



def inputs() -> tuple:

    user_input = int(input('1- Branch/Xfmr \n2- Three Winding Transformer \n3- Shunt \n4- Bus \n5- Gen \nElement Type [1-5]: '))
    user_buses = []

    if user_input == 1:
        user_buses.append(str(input("Enter bus from: ")))
        user_buses.append(str(input("Enter bus to: ")))
    elif user_input == 2:
        user_buses.append(str(input("Enter primary winding bus: ")))
        user_buses.append(str(input("Enter secondary winding bus: ")))
        user_buses.append(str(input("Enter tertiary winding bus: ")))
    elif user_input == 3:
        user_buses.append(str(input("Enter shunt bus number: ")))
    elif user_input == 4:
        user_buses.append(str(input("Enter bus number: ")))
    else:
        user_buses.append(str(input("Enter gen bus number: ")))

    if user_input != 4:
        user_id = str(input('Enter ID of element: '))

    return user_input, user_buses, user_id



def main():

    # User Inputs
    pw_file = filedialog.askopenfilename(title="Select your PowerWorld case", filetypes= [("PowerWorld files", "*.pwb")])
    #input_file = filedialog.askopenfilename("Select input file", filetypes=[("Text Files", "*.txt")])
    #pw_file = '23SSWG_2030_SUM1_U1_Final_10092023_v186.pwb'

    simauto_obj = win32com.client.Dispatch('pwrworld.SimulatorAuto')
    simauto_obj.OpenCase(pw_file)

    err, contingency_elements = simauto_obj.GetParametersMultipleElement('ContingencyElement', ['Contingency', 'ObjectAction', 'Criteria', 'TimeDelay'], '')
    df = pd.DataFrame(contingency_elements)
    df = df.transpose()

    element_type, element_buses, element_id = inputs()

    contingencies = ctg(df, element_type, element_buses, element_id)

    print(" ")
    for i in contingencies:
        print(i)

main()