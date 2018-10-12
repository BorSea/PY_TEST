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

# тут должен быть список, а не строка !!! 

def write_log(string):
   
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        xlsFullPath=logxlsPath + logxlsFile
        wb = openWorkbook(excel, xlsFullPath)
        ws = wb.Worksheets('Лист1')


# строка №4    
# B(2) - номер операции
# C(3) - дата создания архива
# D(4) - имя архива
# E(5) - дата время отправки архива
# F(6) - результат отправки 
        InsRow(ws, string,5,2)

        wb.Close(SaveChanges=True)  
    
    except Exception as e:    
        print(e)

    finally:
        # RELEASES RESOURCES
        ws = None
        wb = None
        excel = None


def InsRow(ws, string, nrow, ncol):

#        ws.Rows.Insert(2) не работает - выход за границы листа
        ws.Cells(nrow, ncol).EntireRow.Insert(Shift=const.xlShiftDown)
        i=0
        for r in string:
# увеличиваем порядковый номер записи  в логе            
            if i == 0:
                ws.Cells(nrow, ncol+i).Value = ws.Cells(nrow, ncol+i).Value + 1
            else:
                ws.Cells(nrow, ncol+i).Value = r

            ws.Cells(nrow, ncol+i).Value = r
            ws.Cells(nrow, ncol+i).Interior.ColorIndex = 6
            i=i+1

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
loglst=[]
b4_num=1
c4_arh_date=''
d4_arh_name=''
e4_arh_send_date=''
f4_act_result='OK'


dt = time.strftime("%Y%m%d_%H%M%S", time.localtime())
zipFilename='WeekZipURS'+dt+'.zip'
arhPAth='C:\\Users\\BorisovSA\\Documents\\Персональное\\send\\'
arhFilename = r'C:\Users\BorisovSA\Documents\Персональное\02-03-04-05-06-07-08-09.2018.xlsx'

# Константы для протокола
logxlsPath="C:\\Users\\BorisovSA\\Documents\\Персональное\\"
logxlsFile="sendmail.xlsx"

    

# if __name__ == '__main__':
# # Занесение информации об отправке в лог
# #loglst=[b4_num, c4_arh_date,d4_arh_name, e4_arh_send_date, f4_act_result]

#     b4_num=1
#     c4_arh_date=dt
#     d4_arh_name=arhFilename
# #08/10/18 0:00    
#     e4_arh_send_date=time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())
#     f4_act_result='OK'


#     loglst.insert(0,b4_num)
#     loglst.insert(1,c4_arh_date)
#     loglst.insert(2,d4_arh_name)
#     loglst.insert(3,e4_arh_send_date)
#     loglst.insert(4,f4_act_result)

#     print('Пишем лог - %s...' % arhFilename + dt)
#     write_log(loglst)




