import tkinter as tk
from . import common_operations
import pandas as pd  # pip install pandas openpyxl xlrd
import PySimpleGUI as sg  # pip install PySimpleGUI
from tkhtmlview import html_parser  # pip install tkhtmlview


def get_preview_layout(size: tuple[int], key: str) -> list[list[sg.Element]]:
    '''Returns the layout to use for the preview tab.'''

    preview_warning_layout = [
        [sg.Push(), sg.Text("This is just a preview to confirm that placeholders are replaced correctly."), sg.Push()],
        [sg.Push(), sg.Text("HTML Styles and Elements might not show correctly."), sg.Push()]
    ]
    column_1_layout = [
        [sg.Text("Recipient E-mail address:")],
        [sg.Text("CC E-mail address:")],
        [sg.Text("BCC E-mail address:")],
        [sg.Text("Subject:")],
        [sg.Text("Attachment(s):")]
    ]
    column_2_layout = [
        [sg.Input(key= "-EMAIL_ADDRESS_PREVIEW-", disabled= True, expand_x= True)],
        [sg.Input(key= "-CC_ADDRESS_PREVIEW-", disabled= True, expand_x= True)],
        [sg.Input(key= "-BCC_ADDRESS_PREVIEW-", disabled= True, expand_x= True)],
        [sg.Input(key= "-SUBJECT_PREVIEW-", disabled= True, expand_x= True)],
        [sg.Input(key= "-ATTACHMENT_PREVIEW-", disabled= True, expand_x= True)]
    ]
    preview_layout = [
        [sg.Frame("WARNING", preview_warning_layout, title_color= "red", expand_x= True)],
        [sg.Text("Row:", tooltip= "Refers to the relative row in excel/CSV file, assuming title is on the first row."), 
            sg.Input(size = (3 , 1), key= "-CURRENT_ROW-", disabled= True), sg.Push(), 
            sg.Button("Show Preview", key= "-SHOW_PREVIEW-"), sg.Push(), 
            sg.Button("Jump to row:", key= "-JUMP_ROW-"), sg.Combo("", readonly= True, size= (3, 1), key= "-ROW_TO_JUMP-")],
        [sg.Frame("E-mail Header Preview", [
            [sg.Column(column_1_layout),
            sg.Column(column_2_layout, expand_x= True) 
            ]
        ], expand_x= True)],
            [sg.Frame("E-mail Body Preview", [[sg.Button(sg.SYMBOL_LEFT_ARROWHEAD, size=(2 , 2), key= "-PREVIOUS-"),
        sg.Multiline(size= size, disabled= True, expand_x= True, expand_y= True, key= key, background_color= "#b3a900"), # sg.theme_background_color()),
        sg.Button(sg.SYMBOL_RIGHT_ARROWHEAD, size= (2, 2), key= "-NEXT-")]], expand_x= True, expand_y= True)]
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


def update_preview_elements(data_df: pd.DataFrame, placeholders: list[str], preview_index: int, values: dict, window: sg.Window) -> None:
    '''Update the preview elements except the message body with the appropriate values from the data dataframe based on the given index.'''

    window.Element("-CURRENT_ROW-").update(preview_index + 2)
    window.Element("-EMAIL_ADDRESS_PREVIEW-").update(common_operations.get_email_address(values["-RECIPIENT_EMAIL_ADDRESS-"], data_df, preview_index))
    if values["-INCLUDE_CC-"] and values["-CC_EMAIL_ADDRESS-"]:
        window.Element("-CC_ADDRESS_PREVIEW-").update(common_operations.get_email_address(values["-CC_EMAIL_ADDRESS-"], data_df, preview_index))
    else:
        window.Element("-CC_ADDRESS_PREVIEW-").update("")
    if values["-INCLUDE_BCC-"] and values["-BCC_EMAIL_ADDRESS-"]:
        window.Element("-BCC_ADDRESS_PREVIEW-").update(common_operations.get_email_address(values["-BCC_EMAIL_ADDRESS-"], data_df, preview_index))
    else:
        window.Element("-BCC_ADDRESS_PREVIEW-").update("")
    window.Element("-SUBJECT_PREVIEW-").update(common_operations.get_subject(values, placeholders, data_df, preview_index))
    window.Element("-ATTACHMENT_PREVIEW-").update(common_operations.get_attachment_filenames(values, data_df, preview_index))   


def show_preview_event(placeholders: list[str], data_df: pd.DataFrame, values: dict, template_text: str, window: sg.Window) -> tuple[int, bool, tk.Widget]:
    '''When the "Show Preview" button is pressed, shows the preview of the first e-mail after replacement of placeholders'''

    preview_index = 0
    if placeholders:
        preview_text = common_operations.replace_placeholders(placeholders, values, data_df, preview_index, template_text) 
    else:
        preview_text = template_text
    preview_element = initialize(window, "-PREVIEW-", preview_text)
    preview_live = True
    update_preview_elements(data_df, placeholders, preview_index, values, window)
    return preview_index, preview_live, preview_element


def next_preview_event(preview_index: int, placeholders: list[str], values: dict, data_df: pd.DataFrame, template_text: str, preview_element: tk.Widget, window: sg.Window) -> int:
    '''Shows the next e-mail preview according to the next dataframe element'''
    
    preview_index += 1
    preview_text = common_operations.replace_placeholders(placeholders, values, data_df, preview_index, template_text)
    set_html(preview_element, preview_text)
    update_preview_elements(data_df, placeholders, preview_index, values, window)
    return preview_index


def previous_preview_event(preview_index: int, placeholders: list[str], values: dict, data_df: pd.DataFrame, template_text: str, preview_element: tk.Widget, window: sg.Window) -> int:
    '''Shows the previous e-mail preview according to the previous dataframe element'''
    
    preview_index -= 1
    preview_text = common_operations.replace_placeholders(placeholders, values, data_df, preview_index, template_text)
    set_html(preview_element, preview_text)
    update_preview_elements(data_df, placeholders, preview_index, values, window)
    return preview_index
    

def jump_to_row_event(values: dict, placeholders: list[str], data_df: pd.DataFrame, template_text: str, preview_element: tk.Widget, window: sg.Window) -> int:
    '''Shows the selected e-mail preview according to the selected dataframe element'''
    
    preview_index = values["-ROW_TO_JUMP-"] - 2
    preview_text = common_operations.replace_placeholders(placeholders, values, data_df, preview_index, template_text)
    set_html(preview_element, preview_text)
    update_preview_elements(data_df, placeholders, preview_index, values, window)
    return preview_index
