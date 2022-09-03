import tkinter as tk
import PySimpleGUI as sg
from tkhtmlview import html_parser


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
        [sg.T("Row:"), sg.I(size = (3,1), key= "-CURRENT_ROW-", disabled= True), sg.Push(), sg.B("Show Preview", key= "-SHOW_PREVIEW-"), sg.Push(), sg.B("Jump to row:", key= "-JUMP_ROW-"), sg.I(size=(3,1), default_text= "2", key= "-ROW_TO_JUMP-")],
        [sg.B(sg.SYMBOL_LEFT_ARROWHEAD, size=(2,2), key= "-PREVIOUS-"),
        sg.Multiline(size= size, disabled= True, expand_x= True, expand_y= True, key= key, background_color= "#b3a900"), #sg.theme_background_color()),
        sg.Button(sg.SYMBOL_RIGHT_ARROWHEAD, size=(2,2), key= "-NEXT-")]
    ]
    return html_layout

def initialize(win: sg.Window, element: sg.Element, html: str) -> tk.Widget:
    '''Initializes the preview Element.'''

    preview_widget = win.Element(element).Widget
    set_html(preview_widget, html)
    return preview_widget

def replace_placeholders(placeholders, values, excel_sheet, preview_index, template_text):
    '''Returns the formatted text after replacing the placeholders for preview'''

    dict = {}
    for placeholder in placeholders:
        column_name = values[("-DATA-", placeholder)]
        df_column = excel_sheet[column_name]
        dict[placeholder] = df_column[preview_index]
    return template_text.format(**dict)


def show_preview_event(placeholders, excel_sheet, values, template_text, window):
    try:
        preview_index = 0
        if placeholders:
            preview_text = replace_placeholders(placeholders, values, excel_sheet, preview_index, template_text) 
            preview_element = initialize(window, "-PREVIEW-", preview_text)
            preview_live = True
            window.Element("-CURRENT_ROW-").update(preview_index + 2)
        return preview_index, preview_live, preview_element
    except Exception as e:
        if e is not TypeError:
            preview_element = window.Element('-PREVIEW-').Widget
            preview_element.delete('1.0', sg.tk.END)
            sg.Window("ERROR", [ 
                [sg.Push(), sg.T("Unable to show preview due to error:"), sg.Push()],
                [sg.Push(), sg.T(f"{str(e)}", text_color= "red"), sg.Push()],
                [sg.Push(), sg.OK(), sg.Push()]
            ], font = ("Arial", 14)).read(close= True)

def next_preview_event(preview_index, placeholders, values, excel_sheet, template_text, preview_element, window):
    preview_index += 1
    preview_text = replace_placeholders(placeholders, values, excel_sheet, preview_index, template_text)
    set_html(preview_element, preview_text)
    window.Element("-CURRENT_ROW-").update(preview_index + 2)
    return preview_index

def previous_preview_event(preview_index, placeholders, values, excel_sheet, template_text, preview_element, window):
    preview_index -= 1
    preview_text = replace_placeholders(placeholders, values, excel_sheet, preview_index, template_text)
    set_html(preview_element, preview_text)
    window.Element("-CURRENT_ROW-").update(preview_index + 2)
    return preview_index