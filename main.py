import PySimpleGUI as sg # pip install PySimpleGUI
import scripts.preview as preview
import scripts.msg_config as msg_config
from scripts.global_constants import THEME, FONT

def main():
    '''Main function'''

    sg.theme(THEME)
    
    message_configuration_layout = msg_config.get_msg_config_layout()
    preview_layout = preview.get_html_layout((40, 20), "-PREVIEW-")
    other_layout = [
        [sg.T("Other")]
    ]
    
    layout = [
        [sg.TabGroup(
            [
                [sg.Tab("Message configuration", message_configuration_layout), 
                sg.Tab("Preview", preview_layout), 
                sg.Tab("E-mail configuration and Send", other_layout)]
            ], expand_x= True, expand_y= True)
        ],
        [sg.B("Reset")]
    ]
    window = sg.Window("Title", layout, font = FONT, return_keyboard_events= True, enable_close_attempted_event= True, resizable= True, finalize= True)

    excel_sheet = None # will be the data dataframe
    placeholders = []
    template_text = ""

    preview_index = 0
    preview_index_max = 0 # will be determined by the amount of rows in dataframe of replacements
    preview_live= False
    preview_element = None
    
    msg_config.expand_scrollable_column(window)

    while True:
        event, values = window.read()
        match event:
            case sg.WINDOW_CLOSE_ATTEMPTED_EVENT | "Exit": # | "Escape:27":
                break
            case "Reset":
                window.close()
                main()

            #----- Message configuration events -----#
            case "-BROWSE_TEMPLATE-":
                msg_config.browse_template_event(window)
            case "-BROWSE_DATA-":
                msg_config.browse_data_event(window)
            case "-GENERATE_PAIRS-":
                try:
                    template_text = msg_config.load_template(values)
                    excel_sheet = msg_config.load_data(values)
                    placeholders = msg_config.generate_pairs_event(placeholders, window, excel_sheet, template_text, values)
                    window.Element("-GENERATE_PAIRS-").update(disabled= True)
                except Exception as e: # not all pairs have been matched
                    sg.Window("ERROR", [
                        [sg.Push(), sg.T("Unable to load file(s):"), sg.Push()],
                        [sg.Push(), sg.T(f"{str(e)}", text_color= "red"), sg.Push()],
                        [sg.Push(), sg.OK(), sg.Push()]
                    ], font = FONT, modal= True).read(close= True)
                #window['-PAIR_COLUMN-'].contents_changed() # appears to not have any effect on the scrollable column, calling after event handling

            #----- preview events -----#
            case "-SHOW_PREVIEW-":
                if excel_sheet is not None:
                    preview_index_max = len(excel_sheet.index) - 1 # minus the title
                    try:
                        preview_index, preview_live, preview_element = preview.show_preview_event(placeholders, excel_sheet, values, template_text, window)
                    except Exception as e: # not all pairs have been matched
                        preview_live = False
                        sg.Window("ERROR", [
                            [sg.Push(), sg.T("Unable to show preview due to error:"), sg.Push()],
                            [sg.Push(), sg.T(f"{str(e)}", text_color= "red"), sg.Push()],
                            [sg.Push(), sg.T(f"Please ensure that all pairs have been matched"), sg.Push()],
                            [sg.Push(), sg.OK(), sg.Push()]
                        ], font = FONT, modal= True).read(close= True)
            case "-NEXT-" | "Right:39":
                if (preview_index < preview_index_max) and preview_live:
                    preview_index = preview.next_preview_event(preview_index, placeholders, values, excel_sheet, template_text, preview_element, window)
            case "-PREVIOUS-" | "Left:37":
                if preview_index > 0 and preview_live:
                    preview_index = preview.previous_preview_event(preview_index, placeholders, values, excel_sheet, template_text, preview_element, window)       
        window['-PAIR_COLUMN-'].contents_changed() # inefficient to call after every event, but does not work in "msg_config.generate_pairs_event" function or after calling it in event handling
    window.close()

if __name__ == "__main__":
    main()
