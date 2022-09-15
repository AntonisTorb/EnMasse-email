import webbrowser
from scripts import attachment_config, email_config_send, msg_config, preview, user_messages, THEME, FONT, FONT_L, FONT_XL, URL_GITHUB
import PySimpleGUI as sg  # pip install PySimpleGUI


def get_info() -> None:
    '''Displays the info window.'''

    info_layout = [
        [sg.Push(), sg.Text("~~EnMasse E-mail~~", font= FONT_XL), sg.Push()],
        [sg.Push(), sg.Text("Send E-mails en masse using a template.", font= FONT_L), sg.Push()],
        [sg.Push(), sg.Text("Version 1.0", font= FONT), sg.Push()],
        [sg.HorizontalSeparator()],
        [sg.Push(), sg.T("Github Repository", key= "-GITHUB_URL-", enable_events= True, tooltip= URL_GITHUB, text_color= "Blue", background_color= "Grey", font= FONT + ("underline", )), sg.Push()],
        [sg.Push(), sg.OK(), sg.Push()]
    ]
    
    info_window = sg.Window("Info", info_layout, enable_close_attempted_event= True, font= FONT, modal= True, icon= "icon.ico")
    while True:
        event, _ = info_window.read()
        match event:
            case sg.WINDOW_CLOSE_ATTEMPTED_EVENT | "OK":
                break
            case "-GITHUB_URL-":
                webbrowser.open(URL_GITHUB)
    info_window.close()


def get_main_layout() -> list[list[sg.Element]]:
    '''Returns the layout to use for the main window.'''

    message_configuration_layout = msg_config.get_msg_config_layout()
    attachment_configuration_layout = attachment_config.get_attachment_config_layout()
    preview_layout = preview.get_preview_layout((30, 10), "-PREVIEW-")
    email_config_send_layout = email_config_send.get_email_config_send_layout()

    layout = [
        [sg.TabGroup(
            [
                [sg.Tab("Message configuration", message_configuration_layout),
                    sg.Tab("Attachment configuration", attachment_configuration_layout),
                    sg.Tab("E-mail configuration and Send", email_config_send_layout),
                    sg.Tab("Preview", preview_layout)
                ]
            ], expand_x= True, expand_y= True)
        ],
        [sg.Button("Reset"), sg.Push(), sg.Button("Info")]
    ]
    return layout


def main():
    '''Main function.'''
    
    sg.theme(THEME)
    #print(sg.theme_text_color())
    
    window = sg.Window("EnMasse E-mail", get_main_layout(), font= FONT, icon= "icon.ico", return_keyboard_events= True, enable_close_attempted_event= True, resizable= True, finalize= True)
    window.set_min_size((1080, 710))

    data_df = None  # Will be the dataframe containing the replacement data.
    placeholders = []
    template_text = ""
    total_emails_to_send = 0  # Will be determined by the amount of rows in data dataframe.
    attachments_valid = False

    preview_index = 0
    preview_live= False
    preview_element = None

    msg_config.expand_scrollable_column(window, "-PAIR_COLUMN-")

    while True:
        event, values = window.read()
        match event:
            case sg.WINDOW_CLOSE_ATTEMPTED_EVENT | sg.WIN_CLOSED: # | "Escape:27":
                break
            case "Reset":
                window.close()
                main()
            case "Info":
                get_info()

            # Message configuration events
            case "-BROWSE_TEMPLATE-":
                msg_config.browse_template_event(window)
            case "-BROWSE_DATA-":
                msg_config.browse_data_event(window)
            case "-GENERATE_PAIRS-":
                if not values["-TEMPLATE_PATH-"] or not values["-DATA_PATH-"]:  # Error handling starts here
                    user_messages.one_line_error_handler("Please ensure you have selected the template and data files.")
                elif not window.Element("-SHEET_NAMES-").Disabled and not values["-SHEET_NAMES-"]:
                    user_messages.one_line_error_handler("Please select the excel sheet.")
                else:
                    try:
                        data_df = msg_config.load_data(values)
                        total_emails_to_send = len(data_df.index)
                        template_text = msg_config.load_template(values)

                        placeholders = msg_config.get_placeholders(placeholders, window, data_df, template_text, values)
                        window.Element("-GENERATE_PAIRS-").update(disabled= True)  # If generate pairs button is pressed again, the resulting pairs have broken element keys for some reason, need to test more. If there is a need to correct something, reset.
                    except Exception as e:
                        user_messages.multiline_error_handler(["Unable to load file(s):", f"{type(e).__name__}: {str(e)}"])
                    #window['-PAIR_COLUMN-'].contents_changed()  # Appears to not have any effect on the scrollable column, calling after event handling.

            # Attachment configuration events
            case "-BROWSE_ATTACHMENT_FILES-":
                attachment_config.browse_attachment_files_event(window)
            case "-BROWSE_ATTACHMENTS_DIRECTORY-":
                attachment_config.browse_attachment_directory_event(window)           
            case "Validate and preview":
                if values["-SEPARATE_ATTACHMENTS-"]:
                    if data_df is None:  # Error handling starts here.
                        user_messages.one_line_error_handler("Please generate the pairs in the first tab.")
                    elif not values["-ATTACHMENTS_DIRECTORY-"]:
                        user_messages.one_line_error_handler("Please specify the dirctory containing the attachment files.")
                    elif not values["-ATTACHMENT_FILENAMES_COLUMN-"]:
                        user_messages.one_line_error_handler("Please select a data column.")
                    else:
                        data_attachments_filenames = attachment_config.get_data_attachments_filenames(values, total_emails_to_send, data_df)
                        directory_attachments_filenames = attachment_config.get_directory_attachments_filenames(values)
                        attachments_valid = attachment_config.attachment_validation(data_attachments_filenames, directory_attachments_filenames)
                        attachment_config.attachment_preview(window, data_attachments_filenames, directory_attachments_filenames)

            # Preview events
            case "-SHOW_PREVIEW-":
                missing_pairs = False
                for placeholder in placeholders:
                    if not values[("-DATA-", placeholder)]:
                        missing_pairs = True
                        break
                if data_df is None:  # Error handling starts here.
                    user_messages.one_line_error_handler("Please generate the pairs in the first tab.")
                elif values["-SUBJECT_ALL-"] and not values["-SUBJECT-"]:
                    user_messages.one_line_error_handler("Please specify the subject.")
                elif values["-SUBJECT_FROM_DATA-"] and not values["-SUBJECT_COLUMN-"]:
                    user_messages.one_line_error_handler("Please specify the subject data column.")
                elif missing_pairs:
                    user_messages.one_line_error_handler("Please ensure that all placeholders have been paired with a data column.")
                elif not values["-RECIPIENT_EMAIL_ADDRESS-"]:
                    user_messages.one_line_error_handler("Please specify the column containing the recipient e-mail address(es).")
                else:
                    try:
                        preview_index, preview_live, preview_element = preview.show_preview_event(placeholders, data_df, values, template_text, window)
                        window.Element("-ROW_TO_JUMP-").update(2, values= [row for row in range(2, total_emails_to_send + 2)])
                    except Exception as e:  # Error while processing the html.
                        preview_live = False
                        user_messages.multiline_error_handler(["Unable to show preview due to error:", f"{type(e).__name__}: {str(e)}"])
            case "-NEXT-" | "Right:39":
                if (preview_index < total_emails_to_send - 1) and preview_live:
                    preview_index = preview.next_preview_event(preview_index, placeholders, values, data_df, template_text, preview_element, window)
            case "-PREVIOUS-" | "Left:37":
                if preview_index > 0 and preview_live:
                    preview_index = preview.previous_preview_event(preview_index, placeholders, values, data_df, template_text, preview_element, window)       
            case "-JUMP_ROW-":
                if preview_live:
                    preview_index = preview.jump_to_row_event(values, placeholders, data_df, template_text, preview_element, window)
            
            # E-mail configuration and send events
            case "-EMAIL_SERVICE-":
                service = values[event]
                email_config_send.set_server_port(service, window)
            case "-PORT-":
                port = values["-PORT-"]
                if len(port) > 3:
                    window.Element("-PORT-").update(port[:3]) 
            case "Setup and Send":
                #print(window.size)  # Testing.
                missing_pairs = False
                for placeholder in placeholders:
                    if not values[("-DATA-", placeholder)]:
                        missing_pairs = True
                        break
                if data_df is None:  # Error handling starts here.
                    user_messages.one_line_error_handler("Please generate the pairs in the first tab.")
                elif values["-SUBJECT_ALL-"] and not values["-SUBJECT-"]:
                    user_messages.one_line_error_handler("PLease specify a subject.")
                elif values["-SUBJECT_FROM_DATA-"] and not values["-SUBJECT_COLUMN-"]:
                    user_messages.one_line_error_handler("Please specify the subject data column.")
                elif missing_pairs:
                    user_messages.one_line_error_handler("Please ensure that all placeholders have been paired with a data column.")
                elif values["-SAME_ATTACHMENTS-"] and not values["-SAME_ATTACHMENT_FILES-"]:
                    user_messages.one_line_error_handler("Please specify the attachment file(s).")
                elif values["-SEPARATE_ATTACHMENTS-"] and not values["-ATTACHMENTS_DIRECTORY-"]:
                    user_messages.one_line_error_handler("Please specify the attachment file directory.")
                elif values["-SEPARATE_ATTACHMENTS-"] and not values["-ATTACHMENT_FILENAMES_COLUMN-"]:
                    user_messages.one_line_error_handler("Please specify the attachment data column.")
                elif values["-SEPARATE_ATTACHMENTS-"] and not attachments_valid:
                    user_messages.one_line_error_handler("Please validate the attachments and ensure that all required files exist in the directory.")               
                elif not values["-SERVER-"] or not values["-PORT-"]:
                    user_messages.one_line_error_handler("Please specify the e-mail server and port.")
                elif not values["-RECIPIENT_EMAIL_ADDRESS-"]:
                    user_messages.one_line_error_handler("Please specify the data column containing the recipient e-mail address(es).")
                elif values["-INCLUDE_CC-"] and not values["-CC_EMAIL_ADDRESS-"]:
                    user_messages.one_line_error_handler("Please specify the data column containing the CC e-mail address(es).")
                elif values["-INCLUDE_BCC-"] and not values["-BCC_EMAIL_ADDRESS-"]:
                    user_messages.one_line_error_handler("Please specify the data column containing the BCC e-mail address(es).")
                elif values["-SET_ALIAS-"] and not values["-ALIAS-"]:
                    user_messages.one_line_error_handler("Please specify your alias.")
                else:
                    email_config_send.setup_and_send_event(data_df, placeholders, template_text, total_emails_to_send, values, window)
        window.Element('-PAIR_COLUMN-').contents_changed()  # Inefficient to call after every event, but does not work in "msg_config.generate_pairs_event" function or after calling it in event handling.
    window.close()

if __name__ == "__main__":
    main()
