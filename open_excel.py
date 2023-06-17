import os
import datetime
import shutil
import xlwings as xw
import win32com.client
def main():
    DirName='.\\'
    FileName='sample.xlsx'
    FilePath=DirName+FileName
    today=datetime.datetime.now()
    num=1
    BkFilePath=DirName+'old\\'+FileName+'_bak{:%y%m%d_%H%M}'.format(today,num)
    shutil.copy(FilePath, BkFilePath)
    wb=xw.Book(r'C:\\Users\\user\\git\\excel_vba\\test02.xlsm')
    macro=wb.macro("open_excel")
    macro(FilePath)
if __name__ == "__main__":
    main()