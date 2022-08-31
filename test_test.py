import openpyxl as xl
import pandas as pd

from openpyxl import load_workbook
import os
import xlwings as xw
import time


EGM_TEMPLATE = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\Data Templates - EGM - WK 34.xlsx'
player_temp = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\Player_Temp.xlsx'
machine_data = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\EGM - Loyalty Data.txt'
all_data = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\EGM - Loyalty Data.txt'
player_temp_txt = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\Player_temp.txt'
#
# Uncarded_sheet = pd.read_excel(EGM_TEMPLATE, sheet_name='Uncarded Data - Report')
# df_Uncarded = pd.DataFrame(Uncarded_sheet)
# xw.view(df_Uncarded)

with xw.App(visible=False) as app:
    book = xw.Book(EGM_TEMPLATE)

    # Do some stuff e.g.
    time.sleep(2)
    book.save()
    book.close()

#
# os.startfile(EGM_TEMPLATE)
# os.close(EGM_TEMPLATE)
# Source_player_sheet = pd.read_excel(EGM_TEMPLATE, sheet_name='Player Data - Source')
# player_sheet = pd.read_excel(EGM_TEMPLATE, sheet_name='Loyalty Data - Report')
# Uncarded_sheet = pd.read_excel(EGM_TEMPLATE, sheet_name='Uncarded Data - Report')
# sheet = pd.read_excel(EGM_TEMPLATE, sheet_name='Sheet1')
#
# df_source = pd.DataFrame(Source_player_sheet)
# df_player = pd.DataFrame(player_sheet)
# df_Uncarded = pd.DataFrame(Uncarded_sheet)
# df_sheet = pd.DataFrame(sheet)
# print(Uncarded_sheet)
# print(df_sheet)