import time, zipfile, os, sys
# Email
import uuid
import itertools as it
# Excel
#import win32com.client as win32
#from win32com.client import DispatchEx
import pywintypes # for exception


# Подключить внешний файл writelog2.py !!!



# Запись строки в протокол работы
sys.path.append('C:\Python\Env')
from writelog2 import *

def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file))

def send_mail(subject,body_text,sendto,copyto=None,blindcopyto=None,attach=None):
   session = DispatchEx('Lotus.NotesSession')
   session.Initialize('Niniga929!')

   # server_name = 'your/server'
   # db_name = 'your/database.nsf'
	
   server_name = "" #r'Main3\Nikoil\RU'
   #	c:\Notes\Data
   db_name = r'C:\Notes\Data\as_СА.nsf'

   db = session.getDatabase(server_name, db_name)
   if not db.IsOpen:
       try:
           db.Open()
       except pywintypes.com_error:
           print( 'could not open database: {}'.format(db_name) )

   doc = db.CreateDocument()
   doc.ReplaceItemValue("Form","Memo")
   doc.ReplaceItemValue("Subject",subject)

   # assign random uid because sometimes Lotus Notes tries to reuse the same one
   uid = str(uuid.uuid4().hex)
   doc.ReplaceItemValue('UNIVERSALID',uid)

   # "SendTo" MUST be populated otherwise you get this error: 
   # 'No recipient list for Send operation'
   doc.ReplaceItemValue("SendTo", sendto)

   if copyto is not None:
       doc.ReplaceItemValue("CopyTo", copyto)
   if blindcopyto is not None:
       doc.ReplaceItemValue("BlindCopyTo", blindcopyto)

   # body
   body = doc.CreateRichTextItem("Body")
   body.AppendText(body_text)

   # attachment 
   if attach is not None:
       attachment = doc.CreateRichTextItem("Attachment")
       for att in attach:
           attachment.EmbedObject(1454, "", att, "Attachment")

   # save in `Sent` view; default is False
   doc.SaveMessageOnSend = True
   doc.Send(False)


# Основная программа

if __name__ == '__main__':

    debug = 1
    
    dt = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    c4_arh_create_date=time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())

    zipFilename='WeekZipURS'+dt+'.zip'
    arhPAth='C:\\Users\\BorisovSA\\Documents\\Персональное\\send\\'
    arhFilename = r'C:\Users\BorisovSA\Documents\Персональное\02-03-04-05-06-07-08-09-10.2018.xlsx'

    #Константы для протокола
    logxlsPath="C:\\Users\\BorisovSA\\Documents\\Персональное\\"
    logxlsFile="sendmail.xlsx"
        
#Создание архива
    newzip=zipfile.ZipFile(zipFilename,'w',zipfile.ZIP_DEFLATED) 
    print('Создание %s...' % (arhFilename))

    if debug == 0 :
        newzip.write(arhFilename)
        zipdir(arhPAth, newzip)
        newzip.close()

        # Отправка
        subject = "WeekEnd"
        body =  zipFilename 
        sendto = ['serj.borisoff@gmail.com',]
            
        #   files = ['/path/to/a/file.txt',]    
        files = [os.path.join(os.getcwd(),zipFilename),]
        attachment = it.takewhile(lambda x: os.path.exists(x), files)
            
        print('Отправка %s...' % (sendto))
        # send_mail(subject, body, sendto, attach=attachment)
            
    else:
    # Занесение информации об отправке в лог
        loglst=[]

        b4_num=0
        c4_arh_date=dt
        d4_arh_name=zipFilename
    #08/10/18 0:00    
        e4_arh_send_date=time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())
        f4_act_result='OK'

        loglst.insert(0,b4_num)
        loglst.insert(1,c4_arh_create_date)
        loglst.insert(2,d4_arh_name)
        loglst.insert(3,e4_arh_send_date)
        loglst.insert(4,f4_act_result)

        print('Пишем лог - %s...' % arhFilename)
        write_log(loglst)



