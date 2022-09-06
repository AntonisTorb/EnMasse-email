import PySimpleGUI as sg # pip install PySimpleGUI
import scripts.preview as preview
import scripts.msg_config as msg_config
import scripts.email_config_send as email_config_send
import scripts.attachment_config as attachment_config
from scripts.global_constants import THEME, FONT
import scripts.user_messages as user_messages


def main():
    '''Main function'''

    sg.theme(THEME)
    #print(sg.theme_text_color())
    
    message_configuration_layout = msg_config.get_msg_config_layout()
    attachment_configuration_layout = attachment_config.get_attachment_config_layout()
    preview_layout = preview.get_preview_layout((40, 10), "-PREVIEW-")
    email_config_send_layout = email_config_send.get_email_config_send_layout()
    
    layout = [
        [sg.TabGroup(
            [
                [sg.Tab("Message configuration", message_configuration_layout),
                sg.Tab("Attachment configuration", attachment_configuration_layout),
                sg.Tab("Preview", preview_layout), 
                sg.Tab("E-mail configuration and Send", email_config_send_layout)]
            ], expand_x= True, expand_y= True)
        ],
        [sg.B("Reset")]
    ]
    window = sg.Window("Title", layout, font = FONT, return_keyboard_events= True, enable_close_attempted_event= True, resizable= True, finalize= True)
    window.set_min_size((1080, 720))

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
            case sg.WINDOW_CLOSE_ATTEMPTED_EVENT | sg.WIN_CLOSED: # | "Escape:27":
                break
            case "Reset":
                window.close()
                main()

            # Message configuration events
            case "-BROWSE_TEMPLATE-":
                msg_config.browse_template_event(window)
            case "-BROWSE_DATA-":
                msg_config.browse_data_event(window)
            case "-GENERATE_PAIRS-":
                if not values["-TEMPLATE_PATH-"] or not values["-DATA_PATH-"]: # error handling starts here
                    user_messages.one_line_error_handler("Please ensure you have selected the template and data files.")
                elif not window.Element("-SHEET_NAMES-").Disabled and not values["-SHEET_NAMES-"]:
                    user_messages.one_line_error_handler("Please select the excel sheet.")
                else:
                    try:
                        template_text = msg_config.load_template(values)
                        excel_sheet = msg_config.load_data(values)
                        placeholders = msg_config.generate_pairs_event(placeholders, window, excel_sheet, template_text, values)
                        window.Element("-GENERATE_PAIRS-").update(disabled= True)
                    except Exception as e:
                        user_messages.multiline_error_handler(["Unable to load file(s):", f"{str(e)}"])
                    #window['-PAIR_COLUMN-'].contents_changed() # appears to not have any effect on the scrollable column, calling after event handling

            # Attachment configuration events
            case "-BROWSE_ATTACHMENT_FILES-":
                attachment_config.browse_attachment_files_event(window)
            case "-BROWSE_ATTACHMENTS_DIRECTORY-":
                attachment_config.browse_attachment_directory_event(window)           
            case "Validate and preview":
                if excel_sheet is None: # error handling starts here
                    user_messages.one_line_error_handler("Please generate the pairs in the first tab.")
                elif not values["-ATTACHMENTS_DIRECTORY-"]:
                    user_messages.one_line_error_handler("Please select an attachment directory.")
                elif not values["-ATTACHMENT_FILENAMES_COLUMN-"]:
                    user_messages.one_line_error_handler("Please select a data column.")
                else:
                    preview_index_max = len(excel_sheet.index) - 1 # minus the title'
                    data_attachments_filenames = attachment_config.get_data_attachments_filenames(values, preview_index_max, excel_sheet)
                    directory_attachments_filenames = attachment_config.get_directory_attachments_filenames(values)
                    attachment_config.attachment_validation(data_attachments_filenames, directory_attachments_filenames)
                    attachment_config.attachment_preview(window, data_attachments_filenames, directory_attachments_filenames)

            # Preview events
            case "-SHOW_PREVIEW-":
                missing_pairs = False
                for placeholder in placeholders:
                    if not values[("-DATA-", placeholder)]:
                        missing_pairs = True
                        break
                if excel_sheet is None: # error handling starts here
                    user_messages.one_line_error_handler("Please generate the pairs in the first tab.")
                elif values["-SUBJECT_ALL-"] and not values["-SUBJECT-"]:
                    user_messages.one_line_error_handler("The subject is blank.")
                elif values["-SUBJECT_FROM_DATA-"] and not values["-SUBJECT_COLUMN-"]:
                    user_messages.one_line_error_handler("Please select the subject data column.")
                elif missing_pairs:
                    user_messages.one_line_error_handler("Please ensure that all placeholders have been paired with a data column.")
                elif not values["-TARGET_EMAIL_ADDRESS-"]:
                    user_messages.one_line_error_handler("Please specify the column containing the target e-mail address(es).")
                else:
                    try:
                        preview_index, preview_live, preview_element = preview.show_preview_event(placeholders, excel_sheet, values, template_text, window)
                        preview_index_max = len(excel_sheet.index) - 1 # minus the title
                        window.Element("-ROW_TO_JUMP-").update(2, values= [row for row in range(2, preview_index_max + 3)])
                    except Exception as e: # not all pairs have been matched or error while processing the html
                        preview_live = False
                        user_messages.n_line_error_handler(["Unable to show preview due to error:", f"{str(e)}", "Please ensure that all pairs have been matched"])
                        # sg.Window("ERROR", [
                        #     [sg.Push(), sg.Text("Unable to show preview due to error:"), sg.Push()],
                        #     [sg.Push(), sg.Text(f"{str(e)}", text_color= "red"), sg.Push()],
                        #     [sg.Push(), sg.Text("Please ensure that all pairs have been matched"), sg.Push()],
                        #     [sg.Push(), sg.OK(), sg.Push()]
                        # ], font = FONT, modal= True).read(close= True)
            case "-NEXT-" | "Right:39":
                if (preview_index < preview_index_max) and preview_live:
                    preview_index = preview.next_preview_event(preview_index, placeholders, values, excel_sheet, template_text, preview_element, window)
            case "-PREVIOUS-" | "Left:37":
                if preview_index > 0 and preview_live:
                    preview_index = preview.previous_preview_event(preview_index, placeholders, values, excel_sheet, template_text, preview_element, window)       
            case "-JUMP_ROW-":
                if preview_live:
                    preview_index = preview.jump_to_row_event(values, placeholders, excel_sheet, template_text, preview_element, window)
            
            #E-mail configuration and send events
            case "-EMAIL_SERVICE-":
                service = values[event]
                email_config_send.set_server_port(service, window)
            case "-PORT-":
                port = values["-PORT-"]
                if len(port) > 3:
                    window.Element("-PORT-").update(port[:3]) 
            case "Setup and Send":
                print(window.size) # testing
                missing_pairs = False
                for placeholder in placeholders:
                    if not values[("-DATA-", placeholder)]:
                        missing_pairs = True
                        break
                if excel_sheet is None: # error handling starts here
                    user_messages.one_line_error_handler("Please generate the pairs in the first tab.")
                elif values["-SUBJECT_ALL-"] and not values["-SUBJECT-"]:
                    user_messages.one_line_error_handler("The subject is blank.")
                elif values["-SUBJECT_FROM_DATA-"] and not values["-SUBJECT_COLUMN-"]:
                    user_messages.one_line_error_handler("Please select the subject data column.")
                elif missing_pairs:
                    user_messages.one_line_error_handler("Please ensure that all placeholders have been paired with a data column.")
                elif not values["-TARGET_EMAIL_ADDRESS-"]:
                    user_messages.one_line_error_handler("Please specify the column containing the target e-mail address(es).")
                elif values["-SAME_ATTACHMENTS-"] and not values["-SAME_ATTACHMENT_FILES-"]:
                    user_messages.one_line_error_handler("Please specify the attachment file(s).")
                elif values["-SEPARATE_ATTACHMENTS-"] and not values["-ATTACHMENTS_DIRECTORY-"]:
                    user_messages.one_line_error_handler("Please specify the attachment file directory.")
                elif values["-SEPARATE_ATTACHMENTS-"] and not values["-ATTACHMENT_FILENAMES_COLUMN-"]:
                    user_messages.one_line_error_handler("Please specify the attachment data column.")
                elif not values["-SERVER-"] or not values["-PORT-"]:
                    user_messages.one_line_error_handler("Please specify the e-mail server and port.")
                elif values["-SET_ALIAS-"] and not values["-ALIAS-"]:
                    user_messages.one_line_error_handler("Please specify your alias.")
                #else:

                    # for i in range(50):   # testing logs
                    #     print(f"Printing {i = }.")

                    #email_config_send.send_emails()
        window.Element('-PAIR_COLUMN-').contents_changed() # inefficient to call after every event, but does not work in "msg_config.generate_pairs_event" function or after calling it in event handling
    window.close()

if __name__ == "__main__":
    main()
