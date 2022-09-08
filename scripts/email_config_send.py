import PySimpleGUI as sg # pip install PySimpleGUI
from scripts.global_constants import FONT, EMAIL_SERVICES, DEFAULT_EMAIL_SETTINGS, URL_GMAIL
import webbrowser
import scripts.user_messages as user_messages

def get_email_config_send_layout() -> list[list[sg.Element]]:
    '''Returns the layout to use for the e-mail configuration and send tab'''

    email_settings_layout = [
        [sg.Text("Select e-mail service:"), 
            sg.Combo(EMAIL_SERVICES, readonly= True, enable_events= True, size= (10, 1), key= "-EMAIL_SERVICE-")],
        [sg.Text("E-mail server:"), sg.Input(expand_x= True, key= "-SERVER-"), 
            sg.Text("Port:"), sg.Input(size= (5, 1), enable_events= True, key= "-PORT-")],
    ]
    alias_layout = [
        [sg.Checkbox("Set alias for sender:", tooltip= "This will appear as the sender instead of your e-mail address", key= "-SET_ALIAS-"),
            sg.Input(expand_x= True, key= "-ALIAS-")]
    ]
    delay_list = [sec for sec in range(1, 6)] + [sec for sec in range (10, 31, 5)]
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
            reroute_stdout= True, # these 2 should not be needed I think, opened issue #5842 on PySimpleGUI GitHub
            reroute_stderr= True
            )],
        [sg.ProgressBar(1, size= (1, 30), key= "-PROGRESS-", visible= False, expand_x= True)]
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
    '''Creates a window to get user credentials and return them, or returns (None, None) if canceled.'''

    password_visible = False

    credentials_layout = [
        [sg.Text("User e-mail address:"), sg.Input(expand_x= True, key= "-EMAIL_ADDRESS-")],
        [sg.Text("User password:"), sg.Input(expand_x= True, key= "-PASSWORD-", password_char= "*"), 
            sg.Button("Show", size= (5, 1), key= "-SHOW_HIDE_PASS-")],
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
    credentials_window = sg.Window("Set credentials and send e-mails", credentials_window_layout, font= FONT, modal= True)
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
                    user_messages.one_line_error_handler("Please enter both e-mail address and password")

def test_without_sending(mail: int) -> None:
    '''Used to test the e-mail sending loop without sending any e-mails.'''

    # if mail == 3:
    #     raise Exception("testing")
    # else:
    print(f"Sending e-mail {mail} ...")


def send_emails(values, window, total_emails_to_send):

    window.Element("-LOG-").update("")
    window.Element('-PROGRESS-').update(visible= True, current_count= 0, max= total_emails_to_send)

    # email_address, password = set_credentials()
    # if email_address is not None and password is not None: # did not cancel
    #     print("Sending...")
    
    # server = values["-SERVER-"]
    # port = values["-PORT-"]
    # if values["-SET_ALIAS-"]
    #    alias = values["-ALIAS-"]
    # else:
    #    alias = ""
    
    # setup smtp connection

    # read delay option and determine the delay in ms if selected
    total_emails_to_send = 10 # remove this, only for testing!
    count_sent = 0
 
    delay = values["-YES_DELAY-"]
    if delay:
        delay_s = int(values["-DELAY-"])
        delay_ms = delay_s * 1000
        window.Element("Setup and Send").update(disabled= True)
        root = window.TKroot
        for mail in range(1, total_emails_to_send + 1):
            root.after(delay_ms, test_without_sending(mail)) # call seperate function to send e-mails here instead of print
            print(f"Successfully sent e-mail {mail}.")
            count_sent += 1
            window.Element('-PROGRESS-').update(current_count = count_sent, max= total_emails_to_send)
            window.refresh()
        print(f"Finished sending {total_emails_to_send} e-mails!")
        window.Element("Setup and Send").update(disabled= False)
    else:
        window.Element("Setup and Send").update(disabled= True)
        for mail in range(1, total_emails_to_send + 1):
            print(f"sending e-mail {mail} ...") # call seperate function to send e-mails here instead of print
            print(f"Successfully sent e-mail {mail}.")
        print(f"Finished sending {total_emails_to_send} e-mails!")
        window.Element("Setup and Send").update(disabled= False)
    
    # generate emails 
    # send one after the other with or without delay
