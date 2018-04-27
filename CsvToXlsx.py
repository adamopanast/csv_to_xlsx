import os
import glob
import csv
from xlsxwriter.workbook import Workbook
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
dir_path = os.path.dirname(os.path.realpath(__file__))

outputFileName = 'CsvToXlsx_output'
os.makedirs('.\%s'%(outputFileName))

for csvfile in glob.glob(os.path.join('.', '*.csv')):
    currentFile = csvfile[:-4] + '.xlsx'
    workbook = Workbook('.\%s\%s'%(outputFileName,currentFile))
    worksheet = workbook.add_worksheet()
    with open(csvfile, 'rt', encoding='utf8') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    workbook.close()
    wb = excel.Workbooks.Open('%s\%s\%s'%(dir_path,outputFileName,currentFile))
    ws = wb.Worksheets("Sheet1")
    ws.Columns.AutoFit()
    wb.Save()
excel.Application.Quit()
