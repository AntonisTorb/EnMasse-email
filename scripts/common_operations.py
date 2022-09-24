from pathlib import Path
import pandas as pd  # pip install pandas openpyxl xlrd


def get_email_address(column, data_df, index):
    '''Returns the e-mail address according to user input.'''
    
    df_column = data_df[column]
    recipient_email_address = df_column[index]
    return recipient_email_address


def replace_placeholders(data_df: pd.DataFrame, index: int, placeholders: list[str], template_text: str, values: dict) -> str:
    '''Returns the formatted text after replacing the placeholders with the dataframe data.'''

    dict = {}
    for placeholder in placeholders:
        column_name = values[("-DATA-", placeholder)]
        df_column = data_df[column_name]
        dict[placeholder] = df_column[index]
    return template_text.format(**dict)


def get_subject(data_df: pd.DataFrame, index: int, placeholders: list[str], values: dict ) -> None:
    '''Returns the subject according to user input.'''

    subject = ""
    if values["-SUBJECT_ALL-"]:
        subject = values["-SUBJECT-"]
        if values["-PAIR_SUBJECT-"]:
            subject = replace_placeholders(data_df, index, placeholders, subject, values)
    elif values["-SUBJECT_FROM_DATA-"] and values["-SUBJECT_COLUMN-"]:
        column_name = values["-SUBJECT_COLUMN-"]
        df_column = data_df[column_name]
        subject = df_column[index]
    return subject


def get_attachment_paths(data_df: pd.DataFrame, index: int, values: dict) -> list[Path]:
    '''Returns a list containing all the attachment file paths.'''

    attachment_paths = []
    if values["-SAME_ATTACHMENTS-"]:
        attachments = values["-SAME_ATTACHMENT_FILES-"]
        
        if ";" in attachments:  # Splitting on ";" if there are multiple attachment files.
            separate_paths = attachments.split(";")
            for path in separate_paths:
                attachment_paths.append(Path(path))
        else:
            attachment_paths.append(Path(attachments))
    elif values["-SEPARATE_ATTACHMENTS-"]:
        directory_path = Path(values["-ATTACHMENTS_DIRECTORY-"])
        column_name = values["-ATTACHMENT_FILENAMES_COLUMN-"]
        df_column = data_df[column_name]
        filenames = df_column[index]
        
        if "," in filenames:  # Splitting on "," if there are multiple attachment files.
            separate_filenames = filenames.split(",")
            for filename in separate_filenames:
                attachment_paths.append(directory_path / filename)
        else:
            attachment_paths.append(directory_path / filenames)
    return attachment_paths


def get_attachment_filenames(data_df:pd.DataFrame, index: int, values: dict) -> list[str]:
    '''Returns the attachment filename(s) according to user input.'''
    
    filenames = ""
    attachment_paths = get_attachment_paths(data_df, index, values)
    
    if attachment_paths:  # A little inefficient to reverse the work of get_attachment_paths, but less repeated code.
        filenames_list = []
        for path in attachment_paths:
            filenames_list.append(path.name)
        filenames = ",".join(filenames_list)
    return filenames
