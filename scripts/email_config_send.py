import PySimpleGUI as sg # pip install PySimpleGUI
from scripts.global_constants import FONT
import webbrowser

EMAIL_SERVICES = ["Gmail", "Yahoo", "Outlook", "Custom"] 
DEFAULT_EMAIL_SETTINGS = {
    "Gmail": {"Server": "smtp.gmail.com", "Port": 465 },
    "Yahoo": {"Server": "smtp.mail.yahoo.com", "Port": 465 },
    "Outlook": {"Server": "smtp-mail.outlook.com", "Port": 587 },
    "Custom": {"Server": "", "Port": "" }
}

def get_email_config_send_layout() -> list[list[sg.Element]]:
    '''Returns the layout to use for the e-mail configuration and send tab'''

    email_settings_layout = [
        [sg.Text("Select e-mail service:"), sg.Combo(EMAIL_SERVICES, readonly= True, enable_events= True, size= (10,1), key= "-EMAIL_SERVICE-")],
        [sg.Text("E-mail sever:"), sg.Input(expand_x= True, key= "-SERVER-"), sg.T("Port:"), sg.I(size= (5,1), enable_events= True, key= "-PORT-")],
    ]
    alias_layout = [
        [sg.Checkbox("Set alias for sender:", tooltip= "This will appear as the sender instead of your e-mail address", key= "-SET_ALIAS-"), sg.I(expand_x= True, key= "-ALIAS-")]
    ]
    delay_list = [sec for sec in range(1,6)] + [sec for sec in range (10, 31, 5)]
    timing_layout = [
        [sg.Radio("Send all e-mails with no delay.", group_id= "set_delay", default= True, key= "-NO_DELAY-")],
        [sg.Radio("Send all e-mails with ", group_id= "set_delay", key= "-YES_DELAY-"),
        sg.Combo(delay_list, default_value= 1, readonly= True, key= "-DELAY-"),
        sg.Text("second(s) delay.")]
    ]
    send_layout = [
        [sg.Push(), sg.Text("Click to set sender credentials and send the e-mails."), sg.Push()],
        [sg.Push(), sg.Button("Setup and Send"), sg.Push()]
    ]
    log_layout = [
        [sg.Multiline("", autoscroll= True, write_only=True, auto_refresh=True, reroute_stdout= True, background_color= sg.theme_background_color(), text_color=sg.theme_text_color(), expand_x= True, expand_y= True, key= "-LOG-")]
    ]
    email_config_send_layout = [
        [sg.Frame("E-mail settings", email_settings_layout, expand_x=True)],
        [sg.Frame("Alias", alias_layout, expand_x= True)],
        [sg.Frame("Timing", timing_layout, expand_x= True), sg.Frame("Send", send_layout, expand_x= True)],
        [sg.Frame("Log", log_layout, expand_x= True, expand_y= True)]
    ]

    return email_config_send_layout

def set_server_port(service: str, window: sg.Window) -> None:
    '''Sets the default server and port settings for the selected e-mail service'''

    window.Element("-SERVER-").update(DEFAULT_EMAIL_SETTINGS[service]["Server"])
    window.Element("-PORT-").update(DEFAULT_EMAIL_SETTINGS[service]["Port"])

def set_credentials() -> tuple[str, str] | tuple[None, None]:


    URL = "https://support.google.com/accounts/answer/185833"
    password_visible = False

    credentials_layout = [
        [sg.Text("User e-mail address:"), sg.Input(expand_x= True, key= "-EMAIL_ADDRESS-")],
        [sg.Text("User password:"), sg.Input(expand_x= True, key= "-PASSWORD-", password_char= "*"), sg.Button("Show", size= (5, 1), key= "-SHOW_HIDE_PASS-")],
    ]
    credentials_warning_layout = [
        [sg.Push(), sg.Text("If you are using a Gmail account, you will have to use "), sg.Push()],
        [sg.Push(), sg.Text("an application password instead of your normal password."), sg.Push()],
        [sg.Push(), sg.Text("For more information please check the following:"), sg.Push()],
        [sg.Push(), sg.Text("Sign in with App Passwords", tooltip= URL, enable_events= True, key= "-APP_PASS_URL-", text_color= "Blue", background_color= "Grey", font= FONT + ("underline",)), sg.Push()],
    ]
    credentials_window_layout = [
        [sg.Frame("Credentials", credentials_layout, expand_x= True)],
        [sg.Frame("Warning", credentials_warning_layout, expand_x= True)],
        [sg.Push(), sg.Button("Send", key= "-SEND-"), sg.Cancel(), sg.Push()],
    ]
    credentials_window = sg.Window("Set credentials and send e-mails", credentials_window_layout, font= FONT, modal= True)
    while True:
        event, values = credentials_window.read()
        match event:
            case sg.WIN_CLOSED | "Cancel":
                credentials_window.close()
                return (None, None)
                break
            case "-APP_PASS_URL-":
                webbrowser.open(URL)
            case "-SHOW_HIDE_PASS-":
                if password_visible:
                    credentials_window.Element("-PASSWORD-").update(password_char= "*")
                    password_visible = False
                    credentials_window.Element("-SHOW_HIDE_PASS-").update("Show")
                else:
                    credentials_window.Element("-PASSWORD-").update(password_char= "")
                    password_visible = True
                    credentials_window.Element("-SHOW_HIDE_PASS-").update("Hide")
            case "-SEND-":
                email_address = values["-EMAIL_ADDRESS-"]
                password = values["-PASSWORD-"]
                if email_address and password:
                    credentials_window.close()
                    return email_address, password
                else:
                    sg.Window("ERROR", [
                        [sg.T("Please enter both e-mail address and password")],
                        [sg.Push(), sg.OK(), sg.Push()]
                    ], font= FONT, modal= True).read(close= True)
    credentials_window.close()

def send_emails():
    email_address, password = set_credentials()
    if email_address is not None and password is not None:
        print("Sending...")
        # read server and port
        # setup smtp connection
        # read delay option
        # generate emails 
        # send one after the other with or without delay
        # print relative excel/csv row after each corresponding message has been sent
        # maybe add a progress bar if possible