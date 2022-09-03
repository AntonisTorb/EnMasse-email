import PySimpleGUI as sg
import tkinter as tk
from pathlib import Path
import pandas as pd
import re

def get_msg_config_layout():
    template_layout = [
        [sg.T("Choose template:"), sg.I(key = "-TEMPLATE_PATH-", expand_x=True, disabled= True), sg.B("Browse", key= "-BROWSE_TEMPLATE-")],
    ]
    data_layout = [
        [sg.T("Choose Data:"), sg.I(key = "-DATA_PATH-", expand_x=True, disabled= True), sg.B("Browse", key= "-BROWSE_DATA-")],
        [sg.T("Choose Sheet if Excel file:"), sg.Push(), sg.Combo([], size= (25,1),  readonly= True, disabled= True, key= "-SHEET_NAMES-")]
    ]
    subject_layout = [
        [sg.Radio("Enter one subject for all messages:", group_id= "get_subject", default= True, key= "-SUBJECT_ALL-"), sg.I(expand_x= True, key= "-SUBJECT-")],
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

def expand_scrollable_column(win):
    '''Allow for elements in scrollable collumn to expand on window resize'''

    frame_id = win.Element('-PAIR_COLUMN-').Widget.frame_id
    canvas = win.Element('-PAIR_COLUMN-').Widget.canvas
    canvas.bind("<Configure>", lambda event, canvas=canvas, frame_id=frame_id:configure(event, canvas, frame_id))

def browse_template_event(window):
    template_to_load = sg.popup_get_file('',file_types=(("HTML files", "*.html*"),("text files", "*.txt*")), no_window=True)
    if template_to_load:
        window["-TEMPLATE_PATH-"].update(Path(template_to_load)) 

def browse_data_event(window):
    data_to_load = sg.popup_get_file('',file_types=(("Excel files", "*.xls*"),("CSV files", "*.csv*")), no_window=True)
    if data_to_load:
        data_path = Path(data_to_load)
        window.Element("-DATA_PATH-").update(data_path)
        if data_to_load.endswith(".xls", 0, -1) or data_to_load.endswith(".xls"):
            with pd.ExcelFile(data_path) as excel_file:
                sheets = excel_file.sheet_names
            window.Element("-SHEET_NAMES-").update(disabled= False, values= sheets)

def load_template(values):
    template_path = Path(values["-TEMPLATE_PATH-"])
    with open(template_path) as template_file:
        return template_file.read()

def load_data(values):
    data_path = Path(values["-DATA_PATH-"])
    sheet_name = values["-SHEET_NAMES-"]
    if str(data_path).endswith(".xls", 0, -1) or str(data_path).endswith(".xls"): # probably need string here instead of path
        excel_sheet =  pd.read_excel(data_path, sheet_name= sheet_name, header= 0)
        data_columns = excel_sheet.columns.values.tolist()    
    return excel_sheet, data_columns


def generate_pairs_event(placeholders, window, values):
    reg = r"\{(.*?)\}" # regex for: group of any characters inside curly brackets
    
    # removing old pairs
    if placeholders:
        for placeholder in placeholders:
            widget = window.Element(placeholder).Widget
            del window.AllKeysDict[placeholder]
            delete_widget(widget.master)
        placeholders = []
    
    # load template
    template_text = load_template(values)
    
    # load data
    excel_sheet, data_columns = load_data(values)

    # enable the subject from column selection
    window.Element("-SUBJECT_COLUMN-").update(disabled= False, values= data_columns)  
    
    # if there is a column named "Subject", select the appropriate option and update the data column value on the dropdown list
    if "Subject" in data_columns: # this appears to not happen in time for the next statement, so if there are any palceholders in subject for all, they will appear in the pairs. Annoying! need to figure it out. 
        window.Element("-SUBJECT_ALL-").update(False)
        window.Element("-SUBJECT_FROM_DATA-").update(True)
        window.Element("-SUBJECT_COLUMN-").update("Subject")
    else:
        window.Element("-SUBJECT_COLUMN-").update("") 
    
    # if one subject for all messages selected, try to find any placeholders in the subject text
    if values["-SUBJECT_ALL-"]:
        subject_all = values["-SUBJECT-"]
    else: 
        subject_all = ""
    
    # combine the lists of placeholders from subject and from template
    temporary_placeholders = re.findall(reg, subject_all) + re.findall(reg, template_text)

    # only keep unique placeholder values
    for placeholder in temporary_placeholders:
        if placeholder in placeholders:
            continue
        else:
            placeholders.append(placeholder)
    
    # show the pairs of placeholders and data by extending the scrollable column at the end of the window
    if placeholders:
        for placeholder in placeholders:
            window.extend_layout(window['-PAIR_COLUMN-'], new_layout(placeholder, data_columns))
        window['-PAIR_COLUMN-'].contents_changed() # does not work for some reason, calling after event handling    

    return template_text, excel_sheet, placeholders

