# import dependencies
from pathlib import Path
import PySimpleGUI as sg
import pandas as pd

# hourly energy parser function
def parse_hourly(txt_file_path, output_folder):

    #function to divide a number by a 1000
    def divide_by_1000(num):
        return (num / 1000).round(2)

    df = pd.read_csv(txt_file_path, delimiter=';')
    output_filename = Path(txt_file_path).stem
    output_file = Path(output_folder) / f"{output_filename}.xlsx"

    df.columns
    df = df[['Start(Malay Peninsula Standard Time)','Vrms_AN_max',
       'Vrms_BN_max', 'Vrms_CN_max', 'Irms_A_max', 'Irms_B_max', 'Irms_C_max', 'PowerP_A_max',
       'PowerP_B_max', 'PowerP_C_max', 'PowerP_Total_max']]
    # convert time column to datetime
    df['Start(Malay Peninsula Standard Time)'] = pd.to_datetime(df['Start(Malay Peninsula Standard Time)'])

    # group data set by dates, and then by hours, round up by 2 decimal places
    result = df.groupby([df['Start(Malay Peninsula Standard Time)'].dt.date, df['Start(Malay Peninsula Standard Time)'].dt.hour])['Vrms_AN_max',
       'Vrms_BN_max', 'Vrms_CN_max', 'Irms_A_max', 'Irms_B_max', 'Irms_C_max', 'PowerP_A_max',
       'PowerP_B_max', 'PowerP_C_max', 'PowerP_Total_max'].mean().round(2)

    # format the Power column from (W) unit to (kW) unit
    result[['PowerP_A_max',
       'PowerP_B_max', 'PowerP_C_max', 'PowerP_Total_max']] = result[['PowerP_A_max',
       'PowerP_B_max', 'PowerP_C_max', 'PowerP_Total_max']].apply(divide_by_1000)
    
    #append (kW) unit to the power columns
    result = result.rename(columns={'PowerP_A_max': 'PowerP_A_max (kW)', 
                                    'PowerP_B_max': 'PowerP_B_max (kW)', 
                                    'PowerP_C_max': 'PowerP_C_max (kW)',
                                    'PowerP_Total_max': 'PowerP_Total_max (kW)'})

    # convert dataset to excel file
    result.to_excel(output_file, index=True)
    sg.popup_no_titlebar("Done!")
    return 

#check for valid filepath
def is_valid_filepath(filepath):
    if filepath and Path(filepath).exists():
        return True

    sg.popup_error("Filepath not correct")
    return False

# GUI Definition
layout = [
    [sg.Text("Input File (in txt format):"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Text Files", "*.txt*"),))],
    [sg.Text("Select Destination Folder:"), sg.Input(key="-OUT-"), sg.FolderBrowse()],
    [sg.Exit(), sg.Button("Parse Data")]
]

window = sg.Window("Energy Logger Data Parser", layout)

while True:
    event, values = window.read()
    print(event, values)
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break

    if(event == "Parse Data"):
        if is_valid_filepath(values["-IN-"]) and is_valid_filepath(values["-OUT-"]):
            try:
                parse_hourly(txt_file_path=values["-IN-"], output_folder=values["-OUT-"])
            except:
                sg.popup_error("Invalid input file")

window.close()