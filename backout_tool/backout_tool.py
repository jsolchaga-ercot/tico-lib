from tkinter import filedialog
import win32com.client
#from pyrtp.pwd import pwd
import pandas as pd

def get_contingency_elements(simauto_obj) -> dict:
    object_type = 'ContingencyElement'
    param_list = ['Contingency', 'ObjectAction', 'Object', 'Action', 'Criteria', 'CriteriaStatus', 'TimeDelay', 'Persistent', 'Comment']
    filter_name = ''

    err, output = simauto_obj.GetParametersMultipleElement(object_type, param_list, filter_name)
    df_ctg = pd.DataFrame(output).transpose()
    df_ctg.columns = param_list

    ctg_dictionary = {}
    for i in range(len(df_ctg)):
        ctg = df_ctg.iloc[i]['Contingency']
        object_action = df_ctg.iloc[i]['ObjectAction']
        if ctg in ctg_dictionary:
            ctg_dictionary[ctg].append(object_action)
        else:
            ctg_dictionary[ctg] = [object_action]

    return ctg_dictionary


def backout_ctg(dictionary_aux, dictionary_base, path) -> tuple:
    new_ctg = {}
    modified_ctg = {}
    deleted_ctg = {}

    for key, value in dictionary_aux.items():
        if key not in dictionary_base:
            new_ctg[key] = value
        if key in dictionary_base and value != dictionary_base[key]:
            modified_ctg[key] = value

    for key, value in dictionary_base.items():
        if key not in dictionary_aux:
            deleted_ctg[key] = value

    with open(path, "a") as file:
        # New contingencies from aux to delete
        file.write('SCRIPT\n{\n')
        for key in new_ctg.keys():
            file.write(f'Delete(Contingency,"<DEVICE>Contingency \'{key}\' \");\n')
        file.write('{\n\n')

        # Deleted contingencies from aux to add
        file.write('Contingency (Name)\n{\n')
        for key, value in deleted_ctg.items():
            file.write(f'\"{key}\"\n')
            file.write('<SUBDATA CTGElement>\n')
            for i in value:
                file.write(f'    \"{i}\"\n')
            file.write('</SUBDATA>\n\n')
        file.write('}\n\n')
            

    return new_ctg, modified_ctg, deleted_ctg



def main():
    # User Inputs PW Case and Aux File
    pw_file = filedialog.askopenfilename(title="Select your PowerWorld case", filetypes= [("PowerWorld files", "*.pwb")])
    aux_file = filedialog.askopenfilename(title="Select aux file to create backout file", filetypes=[("Aux Files", "*.aux")])
    path = aux_file.split('.')[0] + '_BACKOUT.aux'
    # Source Path... CHange
    sourcePath = "\\\\ercot.com\departments\systemplanning\Software Upgrades\SysFiles\idvTOaux\\"
    # Create PowerWorld COM instance
    simauto_obj = win32com.client.Dispatch('pwrworld.SimulatorAuto')
    simauto_obj.OpenCase(pw_file)

    # Apply aux file to case
    simauto_obj.RunScriptCommand("EnterMode(EDIT);")
    simauto_obj.ProcessAuxFIle(aux_file)
    ctg_dictionary_aux = get_contingency_elements(simauto_obj)
    simauto_obj.RunScriptCommand('DiffCaseSetAsBase')
    simauto_obj.CloseCase()

    # Run Difference Case
    simauto_obj.OpenCase(pw_file)
    ctg_dictionary_base = get_contingency_elements(simauto_obj)
    simauto_obj.ProcessAuxFile(sourcePath + "NetworkModel.aux")
    simauto_obj.ProcessAuxFile(sourcePath + "ChangeTolerances.aux")
    simauto_obj.RunScriptCommand(f"DiffFlowWriteCompleteModel(\"{path}\", YES, YES, YES, YES, PRIMARY, \"Network Model\",,,,\"NO\");")
    
    backout_ctg(ctg_dictionary_aux, ctg_dictionary_base, path)

    print('t')


main()