import pandas as pd
from pathlib import Path

def get_email_address(column, data_df, index):
    '''Returns the e-mail address according to user input.'''
    
    df_column = data_df[column]
    recipient_email_address = df_column[index]
    return recipient_email_address

def replace_placeholders(placeholders: list[str], values: dict, data_df: pd.DataFrame, index: int, template_text: str) -> str:
    '''Returns the formatted text after replacing the placeholders with the dataframe data.'''

    dict = {}
    for placeholder in placeholders:
        column_name = values[("-DATA-", placeholder)]
        df_column = data_df[column_name]
        dict[placeholder] = df_column[index]
    return template_text.format(**dict)

def get_subject(values: dict, placeholders: list[str], data_df: pd.DataFrame, index: int) -> None:
    '''Returns the subject according to user input.'''

    subject = ""
    if values["-SUBJECT_ALL-"]:
        subject = values["-SUBJECT-"]
        if values["-PAIR_SUBJECT-"]:
            subject = replace_placeholders(placeholders, values, data_df, index, subject)
    elif values["-SUBJECT_FROM_DATA-"] and values["-SUBJECT_COLUMN-"]:
        column_name = values["-SUBJECT_COLUMN-"]
        df_column = data_df[column_name]
        subject = df_column[index]
    return subject

def get_attachment_paths(data_df: pd.DataFrame, index: int, values: dict) -> list[Path]:
    '''Returns a list containing all the attachment file paths.'''

    attachment_paths = [] # default for no attachments
    if values["-SAME_ATTACHMENTS-"]:
        attachments_path = values["-SAME_ATTACHMENT_FILES-"]
        if ";" in attachments_path:
            separate_paths = attachments_path.split(";")
            for path in separate_paths:
                attachment_paths.append(Path(path))
        else:
            attachment_paths.append(Path(attachments_path))
    elif values["-SEPARATE_ATTACHMENTS-"]:
        directory_path = Path(values["-ATTACHMENTS_DIRECTORY-"])
        column_name = values["-ATTACHMENT_FILENAMES_COLUMN-"]
        df_column = data_df[column_name]
        filenames = df_column[index]
        if "," in filenames:
            separate_filenames = filenames.split(",")
            for filename in separate_filenames:
                attachment_paths.append(directory_path / filename)
        else:
            attachment_paths.append(directory_path / filenames)
    return attachment_paths

def get_attachment_filenames(values, data_df, index):
    '''Returns the attachment filename(s) according to user input.'''
    
    filenames = ""
    attachment_paths = get_attachment_paths(data_df, index, values)
    if attachment_paths: # a little inefficient to reverse the work of get_attachment_paths, but less repeated code
        filenames_list = []
        for path in attachment_paths:
            filenames_list.append(path.name)
        filenames = ",".join(filenames_list)
    return filenames
