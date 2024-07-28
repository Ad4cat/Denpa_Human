import openpyxl
import pandas as pd

MAIN_FILE = 'Denpa_ArenaData.xlsx'

def add_data(num):
    write = openpyxl.load_workbook(MAIN_FILE)["sheet1"]
    ws = write.active
    for i in range(num):
        ws.cell(row=i, column=1, value=i)


    write.save(MAIN_FILE)