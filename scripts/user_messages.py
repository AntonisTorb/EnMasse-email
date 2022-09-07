import PySimpleGUI as sg
from scripts.global_constants import FONT

def one_line_error_handler(text: str) -> None:
    '''Displays window with a short error message.'''

    sg.Window("ERROR!", [
        [sg.Push(), sg.Text(text), sg.Push()],
        [sg.Push(), sg.OK(), sg.Push()]
    ], font= FONT, modal= True).read(close= True)

def multiline_error_handler(string_list: list[str]) -> None:
    '''Displays window with a longer error message.'''

    sg.Window("ERROR!", [
        [[sg.Push(), sg.Text(text), sg.Push()] for text in string_list],
        [sg.Push(), sg.OK(), sg.Push()]
    ], font= FONT, modal= True).read(close= True)

def one_line_warning_handler(text: str) -> None:
    '''Displays window with a short warning message.'''

    sg.Window("Warning", [
        [sg.Push(), sg.Text(text), sg.Push()],
        [sg.Push(), sg.OK(), sg.Push()]
    ], font= FONT, modal= True).read(close= True)

def multiline_warning_handler(string_list: list[str]) -> None:
    '''Displays window with a longer warning message.'''

    sg.Window("Warning", [
        [[sg.Push(), sg.Text(text), sg.Push()] for text in string_list],
        [sg.Push(), sg.OK(), sg.Push()]
    ], font= FONT, modal= True).read(close= True)

def operation_successful(text: str) -> None:
    '''Displays window when an operation completes successfully.'''

    sg.Window("Success", [
            [sg.Push(), sg.Text(text), sg.Push()],
            [sg.Push(), sg.OK(), sg.Push()]
    ], font= FONT, modal= True).read(close= True)
