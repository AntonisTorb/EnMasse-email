from email.message import EmailMessage
from email.utils import formataddr
import mimetypes
from pathlib import Path
import smtplib
import webbrowser
from .global_constants import FONT, EMAIL_SERVICES, DEFAULT_EMAIL_SETTINGS, URL_GMAIL, SHOW, HIDE
from . import common_operations
from . import user_messages
import pandas as pd  # pip install pandas openpyxl xlrd
import PySimpleGUI as sg  # pip install PySimpleGUI


def get_email_config_send_layout() -> list[list[sg.Element]]:
    '''Returns the layout to use for the e-mail configuration and send tab.'''

    email_settings_layout = [
        [sg.Text("Select e-mail service:"), 
            sg.Combo(EMAIL_SERVICES, default_value= EMAIL_SERVICES[0], readonly = True, enable_events = True, size= (10, 1), key= "-EMAIL_SERVICE-")],
        [sg.Text("E-mail server:"), sg.Input(expand_x= True, key= "-SERVER-"), 
            sg.Text("Port:"), sg.Input(size= (5, 1), enable_events= True, key= "-PORT-")],
    ]
    recipients_layout = [
        [sg.Text("Column of recipient e-mail addresses:"), sg.Push(), 
            sg.Combo([], size= (25, 1), key= "-RECIPIENT_EMAIL_ADDRESS-", readonly= True, disabled= True)],
        [sg.Checkbox("Include CC e-mail addresses from column:", key= "-INCLUDE_CC-"), sg.Push(), 
            sg.Combo([], size= (25, 1), key= "-CC_EMAIL_ADDRESS-", readonly= True, disabled= True)],
        [sg.Checkbox("Include BCC e-mail addresses from column:", key= "-INCLUDE_BCC-"), sg.Push(), 
            sg.Combo([], size= (25, 1), key= "-BCC_EMAIL_ADDRESS-", readonly= True, disabled= True)]
    ]
    alias_layout = [
        [sg.Checkbox("Set alias for sender:", tooltip= "This will appear as the sender instead of your e-mail address", key= "-SET_ALIAS-"),
            sg.Input(expand_x= True, key= "-ALIAS-")]
    ]
    send_layout = [
        [sg.Push(), sg.Text("Click to set sender credentials and send the e-mails."), sg.Push()],
        [sg.Push(), sg.Button("Setup and Send"), sg.Push()]
    ]
    log_layout = [
        [sg.Multiline("", 
            autoscroll= True, 
            write_only=True, 
            auto_refresh=True, 
            disabled= True, 
            background_color= sg.theme_background_color(), 
            text_color=sg.theme_text_color(), 
            expand_x= True, 
            expand_y= True, 
            key= "-LOG-", 
            echo_stdout_stderr= True,
            reroute_stdout= True,
            reroute_stderr= True
            )],
        [sg.ProgressBar(1, size= (1, 30), key= "-PROGRESS-", visible= False, bar_color = "green on white", expand_x= True)]
    ]
    email_config_send_layout = [
        [sg.Frame("E-mail settings", email_settings_layout, expand_x= True)],
        [sg.Frame("Recipients", recipients_layout, expand_x= True)],
        [sg.Frame("Alias", alias_layout, expand_x= True)],
        [sg.Frame("Send", send_layout, expand_x= True)],
        [sg.Frame("Log", log_layout, expand_x= True, expand_y= True)]
    ]

    return email_config_send_layout


def set_server_port(service: str, window: sg.Window) -> None:
    '''Sets the default server and port settings for the selected e-mail service.'''

    window.Element("-SERVER-").update(DEFAULT_EMAIL_SETTINGS[service]["Server"])
    window.Element("-PORT-").update(DEFAULT_EMAIL_SETTINGS[service]["Port"])


def set_credentials() -> tuple[str, str] | tuple[None, None]:
    '''Creates a window to get user credentials and return them, or returns (None, None) if canceled.'''

    password_visible = False

    credentials_layout = [
        [sg.Text("User e-mail address:"), sg.Input(expand_x= True, key= "-EMAIL_ADDRESS-")],
        [sg.Text("User password:"), sg.Input(expand_x= True, key= "-PASSWORD-", password_char= "*"), 
            sg.Button("", image_data= HIDE, size= (5, 1), key= "-SHOW_HIDE_PASS-")],
    ]
    credentials_warning_layout = [
        [sg.Push(), sg.Text("If you are using a Gmail account, you will have to use "), sg.Push()],
        [sg.Push(), sg.Text("an application password instead of your normal password."), sg.Push()],
        [sg.Push(), sg.Text("For more information please check the following:"), sg.Push()],
        [sg.Push(), sg.Text("Sign in with App Passwords", tooltip= URL_GMAIL, enable_events= True, key= "-APP_PASS_URL-", text_color= "Blue", background_color= "Grey", font= FONT + ("underline",)), sg.Push()],
    ]
    credentials_window_layout = [
        [sg.Frame("Credentials", credentials_layout, expand_x= True)],
        [sg.Frame("Warning", credentials_warning_layout, expand_x= True)],
        [sg.Push(), sg.Button("Send", key= "-SEND-"), sg.Cancel(), sg.Push()],
    ]
    credentials_window = sg.Window("Set credentials and send e-mails", credentials_window_layout, font= FONT, icon= "icon.ico", modal= True)
    while True:
        event, values = credentials_window.read()
        match event:
            case sg.WIN_CLOSED | "Cancel":
                credentials_window.close()
                return (None, None)
            case "-APP_PASS_URL-":
                webbrowser.open(URL_GMAIL)
            case "-SHOW_HIDE_PASS-":
                if password_visible:
                    credentials_window.Element("-PASSWORD-").update(password_char= "*")
                    password_visible = False
                    credentials_window.Element("-SHOW_HIDE_PASS-").update(image_data= HIDE)
                else:
                    credentials_window.Element("-PASSWORD-").update(password_char= "")
                    password_visible = True
                    credentials_window.Element("-SHOW_HIDE_PASS-").update(image_data= SHOW)
            case "-SEND-":
                email_address = values["-EMAIL_ADDRESS-"]
                password = values["-PASSWORD-"]
                if email_address and password:
                    credentials_window.close()
                    return email_address, password
                else:
                    user_messages.one_line_error_handler("Please enter both e-mail address and password")


def get_email_components(data_df: pd.DataFrame, placeholders: list[str], mail_index: int, template_text: str, values: dict) -> tuple[str, str | None, str | None, str, list[Path], str]:
    '''Returns the recipient e-mail address, CC e-mail address (or None), BCC e-mail address (or None), the e-mail subject,
    the attachment paths in a list, and the e-mail body text, in order to create the e-mail message.
    '''

    recipient_email_address = common_operations.get_email_address(values["-RECIPIENT_EMAIL_ADDRESS-"], data_df, mail_index)
    
    if values["-INCLUDE_CC-"] and values["-CC_EMAIL_ADDRESS-"]:
        cc_email_address = common_operations.get_email_address(values["-CC_EMAIL_ADDRESS-"], data_df, mail_index)
    else:
        cc_email_address = None
    if values["-INCLUDE_BCC-"] and values["-BCC_EMAIL_ADDRESS-"]:
        bcc_email_address = common_operations.get_email_address(values["-BCC_EMAIL_ADDRESS-"], data_df, mail_index)
    else:
        bcc_email_address = None
    subject = common_operations.get_subject(values, placeholders, data_df, mail_index)
    if not values["-NO_ATTACHMENTS-"]:
        attachments_paths = common_operations.get_attachment_paths(data_df, mail_index, values)
    else:
        attachments_paths = []
    if placeholders:
        email_body = common_operations.replace_placeholders(placeholders, values, data_df, mail_index, template_text)
    else:
        email_body = template_text
    return recipient_email_address, cc_email_address, bcc_email_address, subject, attachments_paths, email_body


def create_email(alias: str | None, sender_email_address: str, recipient_email_address: str, cc_email_address: str | None, bcc_email_address: str | None, subject: str, attachments_paths: list[Path], email_body: str) -> EmailMessage:
    '''Create an email object according to the user settings and send it via the provided server connection.'''

    message = EmailMessage()
    message["From"] = formataddr((alias, sender_email_address))
    message["To"] = recipient_email_address
    if cc_email_address is not None:
        message["CC"] = cc_email_address
    if bcc_email_address is not None:
        message["BCC"] = bcc_email_address
    message["Subject"] = subject
    
    message.set_content(email_body, subtype= "plain" )
    if email_body.startswith("<html>"):
        message.add_alternative(email_body, subtype= "html")
    
    if attachments_paths:
        for file in attachments_paths:
            content_type, encoding  = mimetypes.guess_type(file)
            if content_type is None or encoding is not None:
                content_type = 'application/octet-stream'
            maintype, subtype = content_type.split("/", 1)
            with open(file, 'rb') as attachment_file:
                message.add_attachment(attachment_file.read(), maintype= maintype, subtype= subtype, filename= file.name)  
    return message


def setup_and_send_event(data_df: pd.DataFrame, placeholders: list[str], template_text: str, total_emails_to_send: int, values: dict, window: sg.Window) -> None:
    '''Set the e-mail account credentials, create the e-mails according to the user settings and send them via the created server connection.'''

    sender_email_address, password = set_credentials()
    if sender_email_address is not None and password is not None:  # Did not cancel.
        window.Element("-LOG-").update("")
        window.Element('-PROGRESS-').update(visible= True, current_count= 0, max= total_emails_to_send)
        
        server = values["-SERVER-"]
        port = values["-PORT-"]
        
        if values["-SET_ALIAS-"]:
            alias = values["-ALIAS-"]
        else:
            alias = None
        
        count_sent = 0
        print("Establishing connection...")
        try:
            with smtplib.SMTP(server, port) as server:
                server.starttls()
                server.login(sender_email_address, password)
                window.Element("Setup and Send").update(disabled = True)
                
                for mail_index in range(total_emails_to_send):
                    try:
                        recipient_email_address, cc_email_address, bcc_email_address, subject, attachments_paths, email_body = get_email_components(data_df, placeholders, mail_index, template_text, values)
                        print(f"Sending e-mail {mail_index + 1} ...")
                        message = create_email(alias, sender_email_address, recipient_email_address, cc_email_address, bcc_email_address, subject, attachments_paths, email_body)
                        server.send_message(message) 
                        print(f"Successfully sent e-mail {mail_index + 1}.")
                        count_sent += 1
                        window.Element('-PROGRESS-').update(current_count = count_sent, max = total_emails_to_send)
                    except Exception as e:
                        print(f"Error while sending email {mail_index + 1}: {type(e).__name__}, {str(e)}")
                        window.Element('-PROGRESS-').update(bar_color = "red on white")
                print(f"Finished sending {count_sent} e-mails.")
                window.Element("Setup and Send").update(disabled = False)
        except Exception as e:
            print(f"Error with server connection: {type(e).__name__}, {str(e)}")
        
        if count_sent == total_emails_to_send:
            user_messages.operation_successful("Successfully sent all e-mails.")
        else:
            user_messages.multiline_error_handler(["There was an error while sending one or more e-mails.", "Please check the log for the error description and location."])
            