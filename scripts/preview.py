import tkinter as tk
import PySimpleGUI as sg # pip install PySimpleGUI
# html_parser has a typo in the tkhtmlview package, so I am importing a local podule with the typo corrected
#from tkhtmlview import html_parser # pip install tkhtmlview
import scripts.html_parser as html_parser
import pandas as pd # pip install pandas openpyxl xlrd
from pathlib import Path


def get_preview_layout(size: tuple[int], key: str) -> list[list[sg.Element]]:
    '''Returns the layout to use for the preview tab.'''

    preview_warning_layout = [
        [sg.Push(), sg.Text("This is just a preview to confirm that placeholders are replaced correctly."), sg.Push()],
        [sg.Push(), sg.Text("HTML Styles and Elements might not show correctly."), sg.Push()]
    ]
    preview_layout = [
        [sg.Frame("WARNING", preview_warning_layout, title_color= "red", expand_x= True)],
        [sg.Text("Row:", tooltip= "Refers to the relative row in excel/CSV file, assuming title is on the first row."), 
            sg.Input(size = (3,1), key= "-CURRENT_ROW-", disabled= True), sg.Push(), 
            sg.Button("Show Preview", key= "-SHOW_PREVIEW-"), sg.Push(), 
            sg.Button("Jump to row:", key= "-JUMP_ROW-"), sg.Combo("", readonly= True, size=(3,1), key= "-ROW_TO_JUMP-")],
        [sg.Frame("E-mail Header Preview", [
            [sg.Text("Recipient E-mail address:"), sg.Input(key= "-EMAIL_ADDRESS_PREVIEW-", disabled= True, expand_x= True)],
            [sg.Text("Subject:", pad= ((0,158),(0,0))), sg.Input(key= "-SUBJECT_PREVIEW-", disabled= True, expand_x= True)],
            [sg.Text("Attachment(s):", pad= ((0,105),(0,0))), sg.Input(key= "-ATTACHMENT_PREVIEW-", disabled= True, expand_x= True)]
        ], expand_x= True)],
            [sg.Frame("E-mail Body Preview", [[sg.Button(sg.SYMBOL_LEFT_ARROWHEAD, size=(2,2), key= "-PREVIOUS-"),
        sg.Multiline(size= size, disabled= True, expand_x= True, expand_y= True, key= key, background_color= "#b3a900"), # sg.theme_background_color()),
        sg.Button(sg.SYMBOL_RIGHT_ARROWHEAD, size=(2,2), key= "-NEXT-")]], expand_x= True, expand_y= True)]
    ]
    return preview_layout

def set_html(widget: tk.Widget, html: str, strip: bool = True) -> None:
    '''Clears a tkinder widget of its contents and add new contents'''

    parser = html_parser.HTMLTextParser()
    prev_state = widget.cget("state")
    widget.config(state=sg.tk.NORMAL)
    widget.delete("1.0", sg.tk.END)
    widget.tag_delete(widget.tag_names)
    parser.w_set_html(widget, html, strip= strip)
    widget.config(state= prev_state)

def initialize(win: sg.Window, element: sg.Element, html: str) -> tk.Widget:
    '''Initializes the preview Element.'''

    preview_widget = win.Element(element).Widget
    set_html(preview_widget, html)
    return preview_widget

def replace_placeholders(placeholders: list[str], values: dict, data_df: pd.DataFrame, preview_index: int, template_text: str) -> str:
    '''Returns the formatted text after replacing the placeholders for preview'''

    dict = {}
    for placeholder in placeholders:
        column_name = values[("-DATA-", placeholder)]
        df_column = data_df[column_name]
        dict[placeholder] = df_column[preview_index]
    return template_text.format(**dict)

def get_recipient_email_address(values, data_df, preview_index):
    '''Returns the recipient e-mail address preview field according to user input'''
    
    recipient_email_address = ""
    column_name = values["-RECIPIENT_EMAIL_ADDRESS-"]
    df_column = data_df[column_name]
    recipient_email_address = df_column[preview_index]
    return recipient_email_address

def get_subject(values: dict, placeholders: list[str], data_df: pd.DataFrame, preview_index: int) -> None:
    '''Returns subject preview field according to user input'''

    subject = ""
    if values["-SUBJECT_ALL-"]:
        subject = values["-SUBJECT-"]
        if values["-PAIR_SUBJECT-"]:
            subject = replace_placeholders(placeholders, values, data_df, preview_index, subject)
    elif values["-SUBJECT_FROM_DATA-"] and values["-SUBJECT_COLUMN-"]:
        column_name = values["-SUBJECT_COLUMN-"]
        df_column = data_df[column_name]
        subject = df_column[preview_index]
    return subject

def get_attachment_filenames(values, data_df, preview_index):
    '''Returns the attachment(s) preview field according to user input'''
    
    filenames = ""
    if values["-NO_ATTACHMENTS-"]:
        pass
    elif values["-SAME_ATTACHMENTS-"]:
        attachments_path = values["-SAME_ATTACHMENT_FILES-"]
        if ";" in attachments_path:
            seperate = []
            seperate_paths = attachments_path.split(";")
            for path in seperate_paths:
                path_obj = Path(path)
                seperate.append(path_obj.name)
            filenames = ",".join(seperate)
        else:
            path_obj = Path(attachments_path)
            filenames = path_obj.name
    elif values["-SEPARATE_ATTACHMENTS-"]:
        column_name = values["-ATTACHMENT_FILENAMES_COLUMN-"]
        df_column = data_df[column_name]
        filenames = df_column[preview_index]
    return filenames
        

def show_preview_event(placeholders: list[str], data_df: pd.DataFrame, values: dict, template_text: str, window: sg.Window) -> tuple[int, bool, tk.Widget]:
    '''When the "Show Preview" button is pressed, shows the preview of the first e-mail after replacement of placeholders'''

    preview_index = 0
    if placeholders:
        preview_text = replace_placeholders(placeholders, values, data_df, preview_index, template_text) 
        preview_element = initialize(window, "-PREVIEW-", preview_text)
        preview_live = True
        window.Element("-CURRENT_ROW-").update(preview_index + 2)
        window.Element("-EMAIL_ADDRESS_PREVIEW-").update(get_recipient_email_address(values, data_df, preview_index))
        window.Element("-SUBJECT_PREVIEW-").update(get_subject(values, placeholders, data_df, preview_index))
        window.Element("-ATTACHMENT_PREVIEW-").update(get_attachment_filenames(values, data_df, preview_index))
    return preview_index, preview_live, preview_element

def next_preview_event(preview_index: int, placeholders: list[str], values: dict, data_df: pd.DataFrame, template_text: str, preview_element: tk.Widget, window: sg.Window) -> int:
    ''' Shows the next e-mail preview according to the next dataframe element'''
    
    preview_index += 1
    preview_text = replace_placeholders(placeholders, values, data_df, preview_index, template_text)
    set_html(preview_element, preview_text)
    window.Element("-CURRENT_ROW-").update(preview_index + 2)
    window.Element("-EMAIL_ADDRESS_PREVIEW-").update(get_recipient_email_address(values, data_df, preview_index))    
    window.Element("-SUBJECT_PREVIEW-").update(get_subject(values, placeholders, data_df, preview_index))
    window.Element("-ATTACHMENT_PREVIEW-").update(get_attachment_filenames(values, data_df, preview_index))
    return preview_index

def previous_preview_event(preview_index: int, placeholders: list[str], values: dict, data_df: pd.DataFrame, template_text: str, preview_element: tk.Widget, window: sg.Window) -> int:
    ''' Shows the previous e-mail preview according to the previous dataframe element'''
    
    preview_index -= 1
    preview_text = replace_placeholders(placeholders, values, data_df, preview_index, template_text)
    set_html(preview_element, preview_text)
    window.Element("-CURRENT_ROW-").update(preview_index + 2)
    window.Element("-EMAIL_ADDRESS_PREVIEW-").update(get_recipient_email_address(values, data_df, preview_index))
    window.Element("-SUBJECT_PREVIEW-").update(get_subject(values, placeholders, data_df, preview_index))
    window.Element("-ATTACHMENT_PREVIEW-").update(get_attachment_filenames(values, data_df, preview_index))
    return preview_index
    
def jump_to_row_event(values: dict, placeholders: list[str], data_df: pd.DataFrame, template_text: str, preview_element: tk.Widget, window: sg.Window) -> int:
    preview_index = values["-ROW_TO_JUMP-"] - 2
    preview_text = replace_placeholders(placeholders, values, data_df, preview_index, template_text)
    set_html(preview_element, preview_text)
    window.Element("-CURRENT_ROW-").update(preview_index + 2)
    window.Element("-EMAIL_ADDRESS_PREVIEW-").update(get_recipient_email_address(values, data_df, preview_index))
    window.Element("-SUBJECT_PREVIEW-").update(get_subject(values, placeholders, data_df, preview_index))
    window.Element("-ATTACHMENT_PREVIEW-").update(get_attachment_filenames(values, data_df, preview_index))
    return preview_index
