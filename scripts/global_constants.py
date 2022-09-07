THEME = "DarkTeal12"#"DarkBlue14"
FONT = ("Arial", 14)
REG = r"\{(.*?)\}" # regex for: group of any characters inside curly brackets
EMAIL_SERVICES = ["Gmail", "Yahoo", "Outlook", "Custom"] 
DEFAULT_EMAIL_SETTINGS = {
    "Gmail": {
        "Server": "smtp.gmail.com", 
        "Port": 465 
    },
    "Yahoo": {
        "Server": "smtp.mail.yahoo.com", 
        "Port": 465 
    },
    "Outlook": {
        "Server": "smtp-mail.outlook.com", 
        "Port": 587 
    },
    "Custom": {
        "Server": "", 
        "Port": ""
    }
}
URL_GMAIL = "https://support.google.com/accounts/answer/185833"