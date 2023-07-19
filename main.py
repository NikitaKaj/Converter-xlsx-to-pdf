import tkinter as tk

window = tk.Tk()

window.title("Excel to PDF")

label = tk.Label(text="Выберите файл", width=100)
label.grid(column=0, row=0)
entry = tk.Entry(width= 50)
entry.grid(column=0,row=1)

button_openfile = tk.Button(window, text="Выбрать файл...")
button_openfile.grid(column=0, row=2)

# button_openfile.pack()
# label.pack()
# entry.pack()

window.geometry('600x400')

window.mainloop()

# # Import Module
# from win32com import client


# # Open Microsoft Excel
# excel = client.DispatchEx("Excel.Application")
# excel.Interactive = False
# excel.Visible = False

# # Read Excel File
# sheets = excel.Workbooks.open('C:\\Users\\Nikita\\Desktop\\Coverter\\Latviesu.xlsx')
# # work_sheets = sheets.Worksheets[0]

# # Convert into PDF File
# sheets.ExportAsFixedFormat(0, 'C:\\Users\\Nikita\\Desktop\\Coverter\\result')
# sheets.Close()




# from openpyxl import load_workbook

# output_file_name = "Latviesu.xlsx"

# wb = load_workbook(output_file_name, data_only=True)
# ws = wb["Sheet1"]

# ws.cell(5, 6).value = "Something" # 6 по буквами, 5 по цифрам

# wb.save(output_file_name)
# wb.close()