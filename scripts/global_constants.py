THEME = "DarkTeal12"#"DarkBlue14"
FONT = ("Arial", 14)
REG = r"\{([^\{\}]*)\}" # regex for: group of any characters inside curly brackets, except for curly brackets.
DEFAULT_EMAIL_SETTINGS = {
    "Custom": {
        "Server": "", 
        "Port": ""
    },
    "Gmail": {
        "Server": "smtp.gmail.com", 
        "Port": 587 
    },
    "Yahoo": {
        "Server": "smtp.mail.yahoo.com", 
        "Port": 465 
    },
    "Outlook": {
        "Server": "smtp-mail.outlook.com", 
        "Port": 587 
    }
}
EMAIL_SERVICES = list(DEFAULT_EMAIL_SETTINGS.keys())
URL_GMAIL = "https://support.google.com/accounts/answer/185833"