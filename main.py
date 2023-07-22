import PySimpleGUI as sg
import os
import win32com.client as win32

def excel_to_pdf(input_excel_file, sheet_name, cell, output_pdf_file):
    excel = win32.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(input_excel_file)
    ws = wb.Worksheets(sheet_name)
    filename = ws.Range(cell).Value.strip()
    if not filename:
        filename = f"Sheet_{sheet_name}"
    output_pdf_file = os.path.join(output_pdf_file, f"{filename}.pdf")
    ws.ExportAsFixedFormat(0, output_pdf_file)
    wb.Close(True)
    excel.Quit()

sg.theme("Default1")

layout = [
    [sg.Text("Выберите файл Excel:")],
    [sg.Input(key="-FILEPATH-", enable_events=True), sg.FileBrowse()],
    [sg.Text("Выберите листы для конвертации:")],
    [sg.Listbox(values=[], size=(30, 6), key="-SHEETS-", enable_events=True, select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
    [sg.Button("Обновить листы", key="-UPDATE_SHEETS-")],
    [sg.Text("Выберите папку для сохранения PDF:")],
    [sg.Input(key="-OUTPUT_DIR-", enable_events=True), sg.FolderBrowse()],
    [sg.Text("Введите ячейку для названия файла PDF:")],
    [sg.Input(key="-CELL-", size=(10, 1))],
    [sg.Button("Конвертировать в PDF", key="-CONVERT-")]
]

window = sg.Window("Excel to PDF Converter", layout, finalize=True)

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break

    if event == "-FILEPATH-":
        file_path = values["-FILEPATH-"]
        if file_path:
            try:
                import pandas as pd

                df = pd.ExcelFile(file_path)
                sheet_names = df.sheet_names
                window["-SHEETS-"].update(sheet_names)
            except ImportError:
                sg.popup("Ошибка: Не найдена библиотека pandas. Установите ее командой 'pip install pandas'.")

    if event == "-UPDATE_SHEETS-":
        file_path = values["-FILEPATH-"]
        if file_path:
            try:
                import pandas as pd

                df = pd.ExcelFile(file_path)
                sheet_names = df.sheet_names
                window["-SHEETS-"].update(sheet_names)
            except ImportError:
                sg.popup("Ошибка: Не найдена библиотека pandas. Установите ее командой 'pip install pandas'.")

    if event == "-CONVERT-":
        file_path = values["-FILEPATH-"]
        output_dir = values["-OUTPUT_DIR-"]
        selected_sheets = values["-SHEETS-"]
        cell = values["-CELL-"]
        if file_path and output_dir and selected_sheets and cell:
            file_path = os.path.abspath(file_path)
            for sheet_name in selected_sheets:
                excel_to_pdf(file_path, sheet_name, cell, output_dir)
            sg.popup("Успех", "Конвертация в PDF завершена!")

window.close()