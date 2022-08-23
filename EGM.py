import os
from datetime import timedelta, datetime
import pandas as pd
import datetime as dt
import shutil
import regex as re


import xlsxwriter


dates=['2022-08-08', '2022-08-09', '2022-08-10', '2022-08-11', '2022-08-12', '2022-08-13', '2022-08-14']
date_list=[]

original=r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\Account'
target = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\Account-test'

# convert date to actual date

for i in range( len(dates)):
    dates[i] = datetime.strptime(dates[i], '%Y-%m-%d')
    dates[i] = pd.to_datetime(dates[i]) + pd.DateOffset(days=1)

date_list = [t.strftime("%Y-%m-%d") for t in dates]


dir_list = os.listdir(original)

# print(dir_list)

for i in range( len(date_list)):
    date_list[i] = date_list[i].replace('-', '')

date_list = ["Accouting_{0}.txt".format(date) for date in date_list]

# print(date_list)


for i in range(len(date_list)):
    original=os.path.join(original, date_list[i])
    target = os.path.join(target, date_list[i])
    if os.path.isfile(original):
        shutil.copy(original, target)
        original=os.path.split(original)
        original= original[0]
        target = os.path.split(target)
        target = target[0]
# print(target)


target_list = os.listdir(target)
# print(target_list)


#clean data
colspecs = (
    (0, 16),
    (16, 29),
    (29, 46),
    (46, 87),
    (87, 120),
    (120, 133),
    (133, 184),
    (184, 206),
    (206, 233),
    (233, 274),
    (274, 287)
)

data_egm_list = []

for filename in os.listdir(target):
    with open(os.path.join(target, filename), 'r',encoding='utf16') as fin:
       data = fin.read().splitlines(True)
    with open(os.path.join(target, filename), 'w',encoding='utf16') as fout:
        fout.writelines(data[2:])
        # fout.writelines(data[:-2])
        # fout.close()


for filename in os.listdir(target):
   df = pd.read_fwf(os.path.join(target, filename), colspecs=colspecs,encoding='utf16')
   df = df.iloc[2:]
   df = df.iloc[:-2]
   # df.to_excel(r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\test.xlsx',index=False)
   data_egm_list.append(df)

print(data_egm_list)
final_egm_data=pd.concat(data_egm_list)
print(final_egm_data)
final_egm_data.to_excel(r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\test.xlsx',index=False)









