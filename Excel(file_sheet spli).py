import pandas as pd
import os
from openpyxl import load_workbook
import xlsxwriter
from shutil import copyfile

file      = input('File Path: ')
extension = os.path.splitext(file)[1]
filename  = os.path.splitext(file)[0]
pth       = os.path.dirname(file)
newfile   = os.path.join(pth, filename + '_2' + extension)
df        = pd.read_excel(file)
pick_col  = input('Select Column of the Excel file: ')
cols      = list(set(df[pick_col].values))

def send_to_file(cols):
    for i in cols:
        df[df[pick_col] == i].to_excel('{}/{}.xlsx'.format(pth, i), sheet_name = i, index = False)
    print('\nCompleted')
    print('Muy bien wey, hasta luego !')
    return

def send_to_sheet(cols):
    copyfile(file, newfile)
    for j in cols:
        writer = pd.ExcelWriter(newfile, engine = 'openpyxl')
        for myname in cols:
            mydf = df.loc[df[pick_col] == myname]
            mydf.to_excel(writer, sheet_name = myname, index = False)
        writer.save()
    print('/nCompleted')
    print('Tres bien Monsieur, a plus tard !')
    return

print('The data will be split based on these values {} and create {} files or sheets based on next selection. If you are ready to proceed, please type "Y" and ENTER, otherwise type "N"'.format(', '.join(cols), len(cols)))
while True:
    x = input('Ready to proceed (Y / N): ').lower()
    if x == 'y':
        while True:
            s = input('Split into different Sheets or File (S / F): ').lower()
            if s == 'f':
                send_to_file(cols)
                break
            elif s == 's':
                send_to_sheet(cols)
                break
            else: continue
        break
    elif x == 'n':
        print('\nThanks for rocking it.')
        break
