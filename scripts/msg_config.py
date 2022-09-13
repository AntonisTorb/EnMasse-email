from pathlib import Path
import re
import tkinter as tk
from scripts.global_constants import REG
import scripts.user_messages as user_messages
import pandas as pd  # pip install pandas openpyxl xlrd
import PySimpleGUI as sg # pip install PySimpleGUI


def get_msg_config_layout() ->list[list[sg.Element]]:
    '''Returns the layout to use for the message configuration tab.'''

    template_layout = [
        [sg.Text("Choose template:"), sg.Input(key = "-TEMPLATE_PATH-", expand_x= True, disabled= True), 
            sg.Button("Browse", key= "-BROWSE_TEMPLATE-")],
    ]
    data_layout = [
        [sg.Text("Choose Data:"), sg.Input(key = "-DATA_PATH-", expand_x= True, disabled= True), 
            sg.Button("Browse", key= "-BROWSE_DATA-")],
        [sg.Text("Choose Sheet if Excel file:"), sg.Push(), sg.Combo([], size= (25, 1),  readonly= True, disabled= True, key= "-SHEET_NAMES-")]
    ]
    subject_layout = [
        [sg.Radio("Enter one subject for all messages:", group_id= "get_subject", default= True, key= "-SUBJECT_ALL-"), 
            sg.Input(expand_x= True, key= "-SUBJECT-")],
        [sg.Checkbox("Include the subject in Placeholder - Data pair generation.", tooltip = "Check this if the subject contains a placeholer.", pad= ((50, 0) , (0, 0)), key= "-PAIR_SUBJECT-")],
        [sg.Radio("Get subject from Data column:", group_id= "get_subject", key= "-SUBJECT_FROM_DATA-"), sg.Push(), 
            sg.Combo([], size= (25, 1),  readonly= True, disabled= True, key= "-SUBJECT_COLUMN-")]
    ]
    load_pairs_layout = [
        [sg.Button("Generate Placeholder - Data pairs", key= "-GENERATE_PAIRS-")]
    ]
    pair_configure_layout= [
        [sg.Column(layout= [[]], expand_x= True, expand_y= True, scrollable= True, vertical_scroll_only= True, key= "-PAIR_COLUMN-")],
    ]
    msg_config_layout = [
        [sg.Frame("Template", template_layout, expand_x= True)],
        [sg.Frame("Data", data_layout, expand_x= True)],
        [sg.Frame("Subject", subject_layout, expand_x= True)],
        [sg.Frame("Pairing", load_pairs_layout, element_justification= "center", expand_x= True)],
        [sg.Frame("Configure Pairs", pair_configure_layout, expand_x= True, expand_y= True)]
    ]
    return msg_config_layout


def delete_widget(widget: tk.Widget) -> None:
    '''Deletes a tkinder widget and its children widgets.'''

    children = list(widget.children.values())
    for child in children:
        delete_widget(child)
    widget.pack_forget()
    widget.destroy()
    del widget


def new_layout(placeholder_name: str, data_name: list[str]) -> list[list[sg.Element]]:
    '''Returns a new row to be inserted into a PySimpleGUI Column Element.'''

    return [
        [sg.Frame("", [
            [sg.Text("Placeholder: "), 
                sg.Input(placeholder_name, readonly= True, key=("-PLACEHOLDER-", placeholder_name), expand_x= True), 
                sg.Text("Data Column:"), sg.Combo(data_name, size= (25, 1), key=("-DATA-", placeholder_name), readonly= True)
            ]
        ], expand_x= True, key= placeholder_name)]
    ]


def configure(event: tk.Event, canvas: tk.Canvas, frame_id: int) -> None:
    '''Allows elements in a scrollable column to extend upon window resize'''

    canvas.itemconfig(frame_id, width= canvas.winfo_width()) 


def expand_scrollable_column(window: sg.Window, element_key: str) -> None:
    '''Allow for elements in scrollable collumn to expand on window resize'''

    frame_id = window.Element(element_key).Widget.frame_id
    canvas = window.Element(element_key).Widget.canvas
    canvas.bind("<Configure>", lambda event, canvas= canvas, frame_id= frame_id: configure(event, canvas, frame_id))


def browse_template_event(window: sg.Window) -> None:
    '''Updates the template input field with the path to the template file'''

    template_to_load = sg.popup_get_file('', file_types=(("HTML files", "*.html*"), ("text files", "*.txt*")), no_window= True)
    if template_to_load:
        window.Element("-TEMPLATE_PATH-").update(Path(template_to_load)) 


def browse_data_event(window: sg.Window) -> None:
    '''Updates the data input field with the path to the data file. Also updates the excel sheet dropdown if loading excel file'''

    data_to_load = sg.popup_get_file('',file_types=(("Excel files", "*.xls*"), ("CSV files", "*.csv*")), no_window= True)
    if data_to_load:
        data_path = Path(data_to_load)
        window.Element("-DATA_PATH-").update(data_path)
        if data_to_load.endswith(".xls", 0, -1) or data_to_load.endswith(".xls"): #.xls or .xls* files
            with pd.ExcelFile(data_path) as excel_file:
                sheets = excel_file.sheet_names
            window.Element("-SHEET_NAMES-").update(disabled= False, values= sheets)
        else:
            window.Element("-SHEET_NAMES-").update(disabled= True, values= "")


def load_template(values: dict) -> str:
    '''Returns the template file based on the template file path'''

    template_path = Path(values["-TEMPLATE_PATH-"])
    with open(template_path) as template_file:
        return template_file.read()


def find_problems_in_data(data_df: pd.DataFrame) -> None:
    '''Finds columns or rows of a dataframe that have a NaN value and displays warning message providing the relative column and row in excel/csv file.'''

    problem_columns = []
    problem_columns_rows = []
    data_columns = data_df.columns.values.tolist()
    
    for data_column in data_columns:  # Problem value in column title.
        if pd.isnull(data_column):
            problem_columns.append(str(data_column))
    
    for column in data_columns:  # Problem value in data rows.
        column_values_to_list = data_df[column].tolist()
        for value in column_values_to_list:
            if pd.isnull(value):
                row = column_values_to_list.index(value)
                problem_columns_rows.append(f"({column}, {row + 2})")
    
    if problem_columns or problem_columns_rows:
        if len(problem_columns) == 1:
            col_str = str(problem_columns[0])
        else:
            col_str = ", ".join(str(problem_columns))
        if len(problem_columns_rows) == 1:
            col_row_str = str(problem_columns_rows[0])
        else:
            col_row_str = ", ".join(problem_columns_rows)
        user_messages.multiline_warning_handler(["Problems found!", f"Column title(s): {col_str}", f"Value(s) at (column, row): {col_row_str}", "Please reset the app if you wish to correct them."])


def load_data(values: dict) -> pd.DataFrame:
    '''Returns the data dataframe based on the data file path and the data file (and sheet, if loading from excel) selection.'''

    data_path = Path(values["-DATA_PATH-"])
    sheet_name = values["-SHEET_NAMES-"]
    if str(data_path).endswith(".xls", 0, -1) or str(data_path).endswith(".xls"):  # Load data from excel file.
        data_df =  pd.read_excel(data_path, sheet_name= sheet_name, header= 0)
        if data_df.empty:
            raise Exception("Selected sheet has no values.")
        find_problems_in_data(data_df)
    elif str(data_path).endswith(".csv"):  # Load data from .csv file.
        data_df = pd.read_csv(data_path, delimiter= ";", header= 0)
        find_problems_in_data(data_df)
    return data_df


def unique_values_list(given_list: list) -> list:
    '''Returns a list with unique values.'''

    returned_list = []
    for item in given_list:
        if item not in returned_list:
            returned_list.append(item)
    return returned_list


def recipient_actions(window: sg.Window, data_columns: list[str]) -> None:
    """Enables and sets values on the Recipient elements based on the data column names."""

    window.Element("-RECIPIENT_EMAIL_ADDRESS-").update(disabled= False, values= data_columns)
    window.Element("-CC_EMAIL_ADDRESS-").update(disabled= False, values= data_columns)
    window.Element("-BCC_EMAIL_ADDRESS-").update(disabled= False, values= data_columns)


def subject_actions(data_columns: list[str], values: dict, window: sg.Window) -> str:
    '''Enables and sets values on the Subject elements based on the subject setting and the data column names.
    If set to have one subject for all e-mails and have placeholders in the subject, returns the input element value, else returns empty string.
    '''

    window.Element("-SUBJECT_COLUMN-").update(disabled= False, values= data_columns)  # Enable the subject from column selection and add the column names as dropdown list values.
    subject_all = ""
    if "Subject" in data_columns and not values["-PAIR_SUBJECT-"]:
        window.Element("-SUBJECT_ALL-").update(False)
        window.Element("-SUBJECT_FROM_DATA-").update(True)
        window.Element("-SUBJECT_COLUMN-").update("Subject")
    elif values["-SUBJECT_ALL-"] and values["-PAIR_SUBJECT-"]: 
        window.Element("-SUBJECT_ALL-").update(True)
        window.Element("-SUBJECT_FROM_DATA-").update(False)
        subject_all = values["-SUBJECT-"]
        window.Element("-SUBJECT_COLUMN-").update("") 
    # else:
    #     window.Element("-SUBJECT_COLUMN-").update("")
    return subject_all


def attachment_actions(data_columns: list[str], window: sg.Window) -> None:
    '''Enables and sets values on the Attachment elements based on the attachment setting and the data column names'''

    window.Element("-ATTACHMENT_FILENAMES_COLUMN-").update(disabled= False, values= data_columns)  # Enable the attachment filenames from column selection and add the column names as dropdown list values.
    if "Attachments" in data_columns:
        window.Element("-NO_ATTACHMENTS-").update(False)
        window.Element("-SAME_ATTACHMENTS-").update(False)
        window.Element("-SEPARATE_ATTACHMENTS-").update(True)
        window.Element("-ATTACHMENT_FILENAMES_COLUMN-").update("Attachments")


def get_placeholders(placeholders: list[str], window: sg.Window, data_df: pd.DataFrame, template_text: str, values: dict) -> list[str]:
    '''Populates the scrollable column element with "Placeholder - Data" pairs generated from the template text, subject, and the data dataframe.'''

    data_columns = data_df.columns.values.tolist()  # Determine list of columns from dataframe.
    recipient_actions(window, data_columns)
    subject_all = subject_actions(data_columns, values, window)  # Determine what subject to use and, if one subject for all, determine if it has placeholders and return them.
    attachment_actions(data_columns, window)
    placeholder_regex = re.compile(REG)
    temporary_placeholders = placeholder_regex.findall(subject_all) + placeholder_regex.findall(template_text)  # Combine the lists of placeholders from subject and from template.
    placeholders = unique_values_list(temporary_placeholders)  # Only keep unique placeholder values.
    if placeholders:  # Extend the layout of the scrollable column with the "Placeholder - Data" pairs.
        for placeholder in placeholders:
            window.extend_layout(window.Element('-PAIR_COLUMN-'), new_layout(placeholder, data_columns))
            if placeholder in data_columns:
                window.Element(("-DATA-", placeholder)).update(placeholder)
        #window['-PAIR_COLUMN-'].contents_changed()  # Appears to not have any effect on the scrollable column, calling after event handling instead.

    return placeholders
