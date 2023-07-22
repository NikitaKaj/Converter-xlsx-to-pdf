import PySimpleGUI as sg
import os
import win32com.client as win32
import pandas as pd
import re

def get_output_filename(sheet_name, cell_value):
    return cell_value.strip() if cell_value else f"Sheet_{sheet_name}"

def is_valid_cell(cell):
    pattern = r"^[A-Za-z]+\d+$"
    return re.match(pattern, cell)

def excel_to_pdf(input_excel_file, sheet_name, cell, output_pdf_file):
    try:
        excel = win32.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(input_excel_file)
        ws = wb.Worksheets(sheet_name)

        if cell and is_valid_cell(cell):
            try:
                cell_value = ws.Range(cell).Value
            except Exception:
                cell_value = None
        else:
            cell_value = None

        filename = get_output_filename(sheet_name, cell_value)
        output_pdf_file = os.path.join(output_pdf_file, f"{filename}.pdf")
        ws.ExportAsFixedFormat(0, output_pdf_file)
        wb.Close(True)
        excel.Quit()
    except Exception as e:
        sg.popup_error(f"Error occurred: {e}")

sg.theme("Default1")

layout = [
    [sg.Text("Select Excel file:")],
    [sg.Input(key="-FILEPATH-", enable_events=True), sg.FileBrowse()],
    [sg.Text("Select sheets to convert:")],
    [sg.Listbox(values=[], size=(30, 6), key="-SHEETS-", enable_events=True, select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
    [sg.Button("Update sheets", key="-UPDATE_SHEETS-")],
    [sg.Text("Select output folder for PDF files:")],
    [sg.Input(key="-OUTPUT_DIR-", enable_events=True), sg.FolderBrowse()],
    [sg.Text("Enter cell for PDF file name (optional):")],
    [sg.Input(key="-CELL-", size=(10, 1))],
    [sg.Button("Convert to PDF", key="-CONVERT-")]
]

window = sg.Window("Excel to PDF Converter by NicKai", layout, finalize=True)

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break

    if event == "-FILEPATH-":
        file_path = values["-FILEPATH-"]
        if file_path:
            try:
                df = pd.ExcelFile(file_path)
                sheet_names = df.sheet_names
                window["-SHEETS-"].update(sheet_names)
            except Exception as e:
                sg.popup_error(f"Error occurred: Filepath")

    if event == "-UPDATE_SHEETS-":
        file_path = values["-FILEPATH-"]
        if file_path:
            try:
                df = pd.ExcelFile(file_path)
                sheet_names = df.sheet_names
                window["-SHEETS-"].update(sheet_names)
            except Exception as e:
                sg.popup_error(f"Error occurred: Update Sheets")

    if event == "-CONVERT-":
        file_path = values["-FILEPATH-"]
        output_dir = values["-OUTPUT_DIR-"]
        selected_sheets = values["-SHEETS-"]
        cell = values["-CELL-"]
        if file_path and output_dir and selected_sheets:
            try:
                file_path = os.path.abspath(file_path)
                for sheet_name in selected_sheets:
                    excel_to_pdf(file_path, sheet_name, cell, output_dir)
                sg.popup("Success", "Conversion to PDF completed!")
            except Exception as e:
                sg.popup_error(f"Error occurred: File Saving")

window.close()