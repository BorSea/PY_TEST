import time, zipfile, os
# Email
import uuid
import itertools as it
# Excel
import win32com.client as win32


from win32com.client import DispatchEx
import pywintypes # for exception

const = win32.constants


# Запись строки в протокол работы
def write_log(string):
   
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        xlsFullPath=logxlsPath + logxlsFile
        wb = openWorkbook(excel, xlsFullPath)
        ws = wb.Worksheets('Лист2')

        InsRow(ws, string)

        wb.Close(SaveChanges=True)  
    
    except Exception as e:    
        print(e)

    finally:
        # RELEASES RESOURCES
        ws = None
        wb = None
        excel = None


def InsRow(ws, string):

#        ws.Rows.Insert(2) не работает - выход за границы листа
        ws.Cells(2, 1).EntireRow.Insert(Shift=const.xlShiftDown)

        ws.Cells(2, 1).Value = string
#        ws.Cells(1, 1).Value = string + ' 123' 
        return (string)



def openWorkbook(xlapp, xlfile):
    try:        
        xlwb = xlapp.Workbooks(xlfile)            
    except Exception as e:
        try:
            xlwb = xlapp.Workbooks.Open(xlfile)
        except Exception as e:
            print(e)
            xlwb = None                    
    return(xlwb)



# Основная программа

debug = 1
dt = time.strftime("%Y%m%d_%H%M%S", time.localtime())
zipFilename='WeekZipURS'+dt+'.zip'
arhPAth='C:\\Users\\BorisovSA\\Documents\\Персональное\\send\\'
arhFilename = r'C:\Users\BorisovSA\Documents\Персональное\02-03-04-05-06-07-08-09.2018.xlsx'

# Константы для протокола
logxlsPath="C:\\Users\\BorisovSA\\Documents\\Персональное\\"
logxlsFile="sendmail.xlsx"

    

if __name__ == '__main__':
# Занесение информации об отправке в лог
    print('Пишем лог - %s...' % arhFilename + dt)
    write_log(arhFilename + dt)




