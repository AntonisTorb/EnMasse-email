from pathlib import Path
import PySimpleGUI as sg # pip install PySimpleGUI
from scripts.global_constants import FONT
import scripts.user_messages as user_messages
import pandas as pd

def get_attachment_config_layout() -> list[list[sg.Element]]:
    '''Returns the layout to use for the attachment configuration tab.'''
    
    select_attachment_layout = [
        [sg.Radio("No attachments", group_id= "set_attachment_option", default= True , key= "-NO_ATTACHMENTS-")],
        [sg.Radio("Same attachment(s) for all e-mails.", group_id= "set_attachment_option", key= "-SAME_ATTACHMENTS-")],
        [sg.Text("Select attachment file(s):", pad= ((50,0),(0,0))), sg.Input(key= "-SAME_ATTACHMENT_FILES-", disabled= True, expand_x= True), sg.Button("Browse", key= "-BROWSE_ATTACHMENT_FILES-")],
        [sg.Radio("Separate attachment for each file.", group_id= "set_attachment_option", key= "-SEPARATE_ATTACHMENTS-")],
        [sg.Text("Select attachments directory:", pad= ((50,0),(0,0))), sg.Input(key= "-ATTACHMENTS_DIRECTORY-", disabled= True, expand_x= True), sg.Button("Browse", key= "-BROWSE_ATTACHMENTS_DIRECTORY-")],
        [sg.Text("Select the Data column that contains the attachment filenames:", pad= ((50,0),(0,0))), sg.Push(), sg.Combo([], size= (25,1),  readonly= True, disabled= True, key= "-ATTACHMENT_FILENAMES_COLUMN-")]
    ]
    filename_preview_layout = [
        [sg.Push(), sg.Text("Validate if all attachments in data column are in the attachment directory and preview the filenames"), sg.Push()],
        [sg.Push(), sg.B("Validate and preview"), sg.Push()],
        [sg.Push(), sg.Text("In Data:"), sg.Push(), sg.Push(), sg.Text("In Directory:"), sg.Push()],
        [sg.Multiline("", write_only= True, auto_refresh= True, background_color= sg.theme_background_color(), text_color= sg.theme_text_color(), expand_x= True, expand_y= True, key= "-ATTACHMENT_FILENAMES_IN_DATA-"),
        sg.Multiline("", write_only= True, auto_refresh= True, background_color= sg.theme_background_color(), text_color= sg.theme_text_color(), expand_x= True, expand_y= True, key= "-ATTACHMENT_FILENAMES_IN_DIRECTORY-")]
    ]
    attachment_config_layout = [
        [sg.Frame("Select attachment(s)", select_attachment_layout, expand_x= True)],
        [sg.Frame("Attachment filename validation and preview", filename_preview_layout, expand_x= True, expand_y= True)]
    ]
    return attachment_config_layout

def browse_attachment_files_event(window: sg.Window) -> None:
    '''Allows the selection of the attachment file(s) upon pressing the relative Browse button and updates the relative input element.'''

    attachments_to_load = sg.popup_get_file('', multiple_files= True, no_window=True)
    if attachments_to_load:
        attachment_paths = ""
        for attachment in attachments_to_load:
            if attachment_paths:
                attachment_paths = f"{attachment_paths};{Path(attachment)}"
            else:
                attachment_paths = Path(attachment)
            window.Element("-SAME_ATTACHMENT_FILES-").update(attachment_paths) 

def browse_attachment_directory_event(window: sg.Window) -> None:
    '''Allows the selection of the attachment directory upon pressing the relative Browse button and updates the relative input element.'''

    attachments_directory = sg.popup_get_folder('', no_window=True)
    if attachments_directory:
        window.Element("-ATTACHMENTS_DIRECTORY-").update(attachments_directory) 

def attachment_preview(window: sg.Window, attachment_file_names: list[str], directory_filenames:list[str]) -> None:
    '''Displays both the data attachment filename(s) and the directory attachment filename(s) in the multiline element.'''
    
    window.Element("-ATTACHMENT_FILENAMES_IN_DATA-").update("") #clear
    window.Element("-ATTACHMENT_FILENAMES_IN_DIRECTORY-").update("")
    for filename in attachment_file_names:
        window.Element("-ATTACHMENT_FILENAMES_IN_DATA-").update(f"{filename}\n", append= True)
    for filename in directory_filenames:
        window.Element("-ATTACHMENT_FILENAMES_IN_DIRECTORY-").update(f"{filename}\n", append= True)

def get_data_attachments_filenames(values: dict, preview_index_max:int , excel_sheet: pd.DataFrame) -> list[str]:
    '''Returns the filenames of the attachment file(s) from the data.'''
    
    attachments = []
    attachment_file_names = []
    column_name = values["-ATTACHMENT_FILENAMES_COLUMN-"]
    for preview_index in range(preview_index_max + 1):      
        df_column = excel_sheet[column_name]
        attachments.append(df_column[preview_index])
    for attachment in attachments:
        if ";" in attachment: # splitting on ;
            seperate = attachment.split(";")
            for filename in seperate:
                if filename not in attachment_file_names: # only keeping unique values
                    attachment_file_names.append(filename)
        else:
            if attachment not in attachment_file_names: # only keeping unique values
                attachment_file_names.append(attachment)
    attachment_file_names.sort()
    return attachment_file_names

def get_directory_attachments_filenames(values):
    '''Returns the filenames of the attachment file(s) in the selected directory.'''

    directory_files = list(Path(values["-ATTACHMENTS_DIRECTORY-"]).glob("*"))
    return [file.name for file in directory_files]

def attachment_validation(data_attachments_filenames: list[str], directory_attachments_filesnames: list[str]) -> None:
    '''Validating if all attachment names from the column exist in the directory, and if therre are any usused files in the directory.'''

    all_found = True
    not_found= []
    all_used = True
    unused = []
    for attachment_filename in data_attachments_filenames:
        if attachment_filename not in directory_attachments_filesnames:
            all_found = False
            not_found.append(attachment_filename)
    if all_found:
        user_messages.operation_successful("All filenames have been located in the specified directory")
    else:
        user_messages.multiline_error_handler(["File(s) not found in the selected directory:", ", ".join(not_found)])
    for attachment_filename in directory_attachments_filesnames:
        if attachment_filename not in data_attachments_filenames:
            all_used = False
            unused.append(attachment_filename)
    if not all_used:
        user_messages.multiline_warning_handler(["There are unused files in the directory:", ", ".join(unused)])

            