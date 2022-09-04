import PySimpleGUI as sg # pip install PySimpleGUI
import tkinter as tk
from pathlib import Path
import pandas as pd # pip install pandas openpyxl xlrd
import re
from scripts.global_constants import REG

def get_msg_config_layout() ->list[list[sg.Element]]:
    '''Returns the layout to use for the message configuration tab'''

    template_layout = [
        [sg.T("Choose template:"), sg.I(key = "-TEMPLATE_PATH-", expand_x=True, disabled= True), sg.B("Browse", key= "-BROWSE_TEMPLATE-")],
    ]
    data_layout = [
        [sg.T("Choose Data:"), sg.I(key = "-DATA_PATH-", expand_x=True, disabled= True), sg.B("Browse", key= "-BROWSE_DATA-")],
        [sg.T("Choose Sheet if Excel file:"), sg.Push(), sg.Combo([], size= (25,1),  readonly= True, disabled= True, key= "-SHEET_NAMES-")]
    ]
    subject_layout = [
        [sg.Radio("Enter one subject for all messages:", group_id= "get_subject", default= True, key= "-SUBJECT_ALL-"), sg.I(expand_x= True, key= "-SUBJECT-")],
        [sg.Checkbox("Include the subject in Placeholder - Data pair generation.", pad= ((50,0),(0,0)), key= "-PAIR_SUBJECT-")],
        [sg.Radio("Get subject from Data column:", group_id= "get_subject", key= "-SUBJECT_FROM_DATA-"), sg.Push(), sg.Combo([], size= (25,1),  readonly= True, disabled= True, key= "-SUBJECT_COLUMN-")]
    ]
    load_pairs_layout = [
        [sg.B("Generate Placeholder - Data pairs", key= "-GENERATE_PAIRS-")]
    ]
    pair_configure_layout= [
        [sg.Column(layout= [[]], expand_x= True, expand_y= True, scrollable= True, vertical_scroll_only= True, key= "-PAIR_COLUMN-")],
    ]
    msg_config_layout = [
        [sg.Frame("Template", template_layout, expand_x=True)],
        [sg.Frame("Data", data_layout, expand_x=True)],
        [sg.Frame("Subject", subject_layout, expand_x=True)],
        [sg.Frame("Pairing", load_pairs_layout, element_justification= "center", expand_x=True)],
        [sg.Frame("Configure Pairs", pair_configure_layout, expand_x=True, expand_y= True)]
    ]
    return msg_config_layout

def delete_widget(widget: tk.Widget) -> None:
    '''Deletes a tkinder widget and its children widgets'''

    children = list(widget.children.values())
    for child in children:
        delete_widget(child)
    widget.pack_forget()
    widget.destroy()
    del widget

def new_layout(placeholder_name: str, data_name: list[str]) -> list[list[sg.Element]]:
    '''Returns a new row to be inserted into a PySimpleGUI Column Element'''

    return [
        [sg.Frame("",[[sg.T("Placeholder: "), 
        sg.I(placeholder_name, readonly= True, size= (32,1), key=("-PLACEHOLDER-", placeholder_name), expand_x= True), 
        sg.T("Data Column:"), sg.Combo(data_name, key=("-DATA-", placeholder_name), readonly= True)]], expand_x= True, key= placeholder_name)]
    ]


def configure(event: tk.Event, canvas: tk.Canvas, frame_id: int) -> None:
    '''Allows elements in a scrollable column to extend upon window resize'''

    canvas.itemconfig(frame_id, width=canvas.winfo_width()) 

def expand_scrollable_column(window: sg.Window) -> None:
    '''Allow for elements in scrollable collumn to expand on window resize'''

    frame_id = window.Element('-PAIR_COLUMN-').Widget.frame_id
    canvas = window.Element('-PAIR_COLUMN-').Widget.canvas
    canvas.bind("<Configure>", lambda event, canvas=canvas, frame_id=frame_id:configure(event, canvas, frame_id))

def browse_template_event(window: sg.Window) -> None:
    '''Updates the template input field with the path to the template file'''

    template_to_load = sg.popup_get_file('',file_types=(("HTML files", "*.html*"),("text files", "*.txt*")), no_window=True)
    if template_to_load:
        window["-TEMPLATE_PATH-"].update(Path(template_to_load)) 

def browse_data_event(window: sg.Window) -> None:
    '''Updates the data input field with the path to the data file. Also updates the excel sheet dropdown if loading excel file'''

    data_to_load = sg.popup_get_file('',file_types=(("Excel files", "*.xls*"),("CSV files", "*.csv*")), no_window=True)
    if data_to_load:
        data_path = Path(data_to_load)
        window.Element("-DATA_PATH-").update(data_path)
        if data_to_load.endswith(".xls", 0, -1) or data_to_load.endswith(".xls"): #.xls or .xls* files
            with pd.ExcelFile(data_path) as excel_file:
                sheets = excel_file.sheet_names
            window.Element("-SHEET_NAMES-").update(disabled= False, values= sheets)

def load_template(values: dict) -> str:
    '''Returns the template file based on the template file path'''

    template_path = Path(values["-TEMPLATE_PATH-"])
    with open(template_path) as template_file:
        return template_file.read()

def load_data(values: dict) -> pd.DataFrame:
    '''Returns the data dataframe based on the data file path and the excel sheet selection (if applicable)'''

    data_path = Path(values["-DATA_PATH-"])
    sheet_name = values["-SHEET_NAMES-"]
    if str(data_path).endswith(".xls", 0, -1) or str(data_path).endswith(".xls"):
        excel_sheet =  pd.read_excel(data_path, sheet_name= sheet_name, header= 0)
    return excel_sheet

def unique_values_list(given_list: list) ->list:
    '''Returns a list with unique values'''

    returned_list = []
    for item in given_list:
        if item not in returned_list:
            returned_list.append(item)
    return returned_list

def subject_actions(data_columns: list[str], values: dict, window: sg.Window) -> str:
    '''Performs action based on the type of subject selected by the user and returns the subject if subject for all is selected, else returns an empty string'''

    subject_all = ""
    if "Subject" in data_columns and not values["-PAIR_SUBJECT-"]:
        window.Element("-SUBJECT_ALL-").update(False)
        window.Element("-SUBJECT_FROM_DATA-").update(True)
        window.Element("-SUBJECT_COLUMN-").update("Subject")
    elif values["-SUBJECT_ALL-"] and values["-PAIR_SUBJECT-"]: 
        subject_all = values["-SUBJECT-"]
        window.Element("-SUBJECT_COLUMN-").update("") 
    else:
        window.Element("-SUBJECT_COLUMN-").update("")
    return subject_all

def generate_pairs_event(placeholders: list[str], window:sg.Window, excel_sheet: pd.DataFrame, template_text: str, values: dict) -> list[str]:
    '''Populates the scrollable column element with "Placeholder - Data" pairs generated from the template text and the data dataframe.'''

    if placeholders: # removing old pairs from layout if any exist and resetting the placeholders list to empty
        for placeholder in placeholders:
            widget = window.Element(placeholder).Widget
            del window.AllKeysDict[placeholder]
            delete_widget(widget.master)
        placeholders = []
    data_columns = excel_sheet.columns.values.tolist() # determine list of columns from dataframe  
    window.Element("-SUBJECT_COLUMN-").update(disabled= False, values= data_columns) # enable the subject from column selection and add the column names as dropdown list values
    subject_all = subject_actions(data_columns, values, window) # determine what subject to use and, if one subject for all, determine if it has placeholders and return them 
    temporary_placeholders = re.findall(REG, subject_all) + re.findall(REG, template_text) # combine the lists of placeholders from subject and from template
    placeholders = unique_values_list(temporary_placeholders) # only keep unique placeholder values
    if placeholders: # extend the layout of the scrollable column with the "Placeholder - Data" pairs
        for placeholder in placeholders:
            window.extend_layout(window['-PAIR_COLUMN-'], new_layout(placeholder, data_columns))
        #window['-PAIR_COLUMN-'].contents_changed() # appears to not have any effect on the scrollable column, calling after event handling
    return placeholders
