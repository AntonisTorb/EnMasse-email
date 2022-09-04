import tkinter as tk
import PySimpleGUI as sg # pip install PySimpleGUI
from tkhtmlview import html_parser # pip install tkhtmlview
import pandas as pd # pip install pandas openpyxl xlrd


def set_html(widget: tk.Widget, html: str, strip: bool = True) -> None:
    '''Clears a tkinder widget of its contents and add new contents'''

    parser = html_parser.HTMLTextParser()
    prev_state = widget.cget("state")
    widget.config(state=sg.tk.NORMAL)
    widget.delete("1.0", sg.tk.END)
    widget.tag_delete(widget.tag_names)
    parser.w_set_html(widget, html, strip= strip)
    widget.config(state= prev_state)

def get_html_layout(size: tuple[int], key: str) -> list[list[sg.Element]]:
    '''Returns the layout to use for the preview tab.'''

    warning_layout = [
        [sg.Push(), sg.T("This is just a preview to confirm that placeholders are replaced correctly."), sg.Push()],
        [sg.Push(), sg.T("HTML Styles and Elements might not show correctly."), sg.Push()]
    ]
    html_layout = [
        [sg.Frame("WARNING", warning_layout, title_color= "red", expand_x= True)],
        [sg.T("Row:"), sg.I(size = (3,1), key= "-CURRENT_ROW-", disabled= True), 
        sg.Push(), 
        sg.B("Show Preview", key= "-SHOW_PREVIEW-"), 
        sg.Push(), 
        sg.B("Jump to row:", key= "-JUMP_ROW-"), sg.I(size=(3,1), default_text= "2", key= "-ROW_TO_JUMP-")],
        [sg.Frame("Subject Preview", [[sg.T("Subject:"), sg.I(key= "-SUBJECT_PREVIEW-", disabled= True, background_color= "#b3a900", expand_x= True)]], expand_x= True)],
        [sg.Frame("E-mail Body Preview", [[sg.B(sg.SYMBOL_LEFT_ARROWHEAD, size=(2,2), key= "-PREVIOUS-"),
        sg.Multiline(size= size, disabled= True, expand_x= True, expand_y= True, key= key, background_color= "#b3a900"), # sg.theme_background_color()),
        sg.Button(sg.SYMBOL_RIGHT_ARROWHEAD, size=(2,2), key= "-NEXT-")]], expand_x= True, expand_y= True)]
    ]
    return html_layout

def initialize(win: sg.Window, element: sg.Element, html: str) -> tk.Widget:
    '''Initializes the preview Element.'''

    preview_widget = win.Element(element).Widget
    set_html(preview_widget, html)
    return preview_widget

def replace_placeholders(placeholders: list[str], values: dict, excel_sheet: pd.DataFrame, preview_index: int, template_text: str) -> str:
    '''Returns the formatted text after replacing the placeholders for preview'''

    dict = {}
    for placeholder in placeholders:
        column_name = values[("-DATA-", placeholder)]
        df_column = excel_sheet[column_name]
        dict[placeholder] = df_column[preview_index]
    return template_text.format(**dict)

def preview_subject(values: dict, placeholders: list[str], excel_sheet: pd.DataFrame, preview_index: int, window: sg.Window) -> None:
    '''Updates the subject preview field according to user input'''

    if values["-SUBJECT_ALL-"]:
        subject = values["-SUBJECT-"]
        if values["-PAIR_SUBJECT-"]:
            subject = replace_placeholders(placeholders, values, excel_sheet, preview_index, subject)
        window.Element("-SUBJECT_PREVIEW-").update(subject)
    elif values["-SUBJECT_FROM_DATA-"] and values["-SUBJECT_COLUMN-"]:
        column_name = values["-SUBJECT_COLUMN-"]
        df_column = excel_sheet[column_name]
        subject = df_column[preview_index]
        print(subject)
        window.Element("-SUBJECT_PREVIEW-").update(subject)

def show_preview_event(placeholders: list[str], excel_sheet: pd.DataFrame, values: dict, template_text: str, window: sg.Window) -> tuple[int, bool, tk.Widget]:
    '''When the "Show Preview" button is pressed, shows the preview of the first e-mail after replacement of placeholders'''

    preview_index = 0
    if placeholders:
        preview_text = replace_placeholders(placeholders, values, excel_sheet, preview_index, template_text) 
        preview_element = initialize(window, "-PREVIEW-", preview_text)
        preview_live = True
        window.Element("-CURRENT_ROW-").update(preview_index + 2)
        preview_subject(values, placeholders, excel_sheet, preview_index, window)
    return preview_index, preview_live, preview_element

def next_preview_event(preview_index: int, placeholders: list[str], values: dict, excel_sheet: pd.DataFrame, template_text: str, preview_element: tk.Widget, window: sg.Window) -> int:
    ''' Shows the next e-mail preview according to the next dataframe element'''
    
    preview_index += 1
    preview_text = replace_placeholders(placeholders, values, excel_sheet, preview_index, template_text)
    set_html(preview_element, preview_text)
    window.Element("-CURRENT_ROW-").update(preview_index + 2)
    preview_subject(values, placeholders, excel_sheet, preview_index, window)
    return preview_index

def previous_preview_event(preview_index: int, placeholders: list[str], values: dict, excel_sheet: pd.DataFrame, template_text: str, preview_element: tk.Widget, window: sg.Window) -> int:
    ''' Shows the previous e-mail preview according to the previous dataframe element'''
    
    preview_index -= 1
    preview_text = replace_placeholders(placeholders, values, excel_sheet, preview_index, template_text)
    set_html(preview_element, preview_text)
    window.Element("-CURRENT_ROW-").update(preview_index + 2)
    preview_subject(values, placeholders, excel_sheet, preview_index, window)
    return preview_index
    