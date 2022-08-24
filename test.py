import openpyxl as xl
import pandas as pd
data = pd.read_excel(r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\Data Templates - EGM - WK 34.xlsx',sheet_name='Floor Analysis - Data')
df = pd.DataFrame(data)
print(df)