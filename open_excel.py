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
    BkFilePath=DirName+'backup_folder\\bak{:%y%m%d_%H%M%S}_'.format(today,num)+FileName
    shutil.copy(FilePath, BkFilePath)
    wb=xw.Book(r'C:\\Users\\user\\git\\excel_vba\\test02.xlsm')
    macro=wb.macro("open_excel")
    macro(FilePath)
if __name__ == "__main__":
    main()