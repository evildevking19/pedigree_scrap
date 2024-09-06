import sys, base64
from multiprocessing import Process
from constants import getGoogleService, createFileWith, ORDER_DIR_NAME

import thread1 as t1
import thread2 as t2
import thread3 as t3

import dam_thread1 as dt1
import dam_thread2 as dt2

class Unbuffered(object):
   def __init__(self, stream):
       self.stream = stream
   def write(self, data):
       self.stream.write(data)
       self.stream.flush()
   def writelines(self, datas):
       self.stream.writelines(datas)
       self.stream.flush()
   def __getattr__(self, attr):
       return getattr(self.stream, attr)

def run(sheetId, sheetName):
    sys.stdout = Unbuffered(sys.stdout)

    print("Preparing to run script...")
    pdf_cnt = 0
    service = getGoogleService('gmail', 'v1')
    results = service.users().messages().list(userId='me', labelIds=['INBOX'], q="from:noreply@aqha.org").execute()
    messages = results.get('messages')
    if messages:
        for message in messages:
            msg_id = message['id']
            msg_body = service.users().messages().get(userId='me', id=msg_id).execute()
            try:
                parts = msg_body['payload']['parts']
                for part in parts:
                    if part['filename']:
                        if 'data' in part['body']:
                            data = part['body']['data']
                        else:
                            att_id = part['body']['attachmentId']
                            att = service.users().messages().attachments().get(userId='me', messageId=msg_id, id=att_id).execute()
                            data = att['data']
                        file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                        filename = part['filename']
                        
                        createFileWith(ORDER_DIR_NAME + "/" + filename, file_data, 'wb')
                        print(filename)
                        pdf_cnt += 1
                service.users().messages().delete(userId='me', id=msg_id).execute()
            except: continue

    print("Preparing done.")
    proc1 = Process(target=t1.start, args=[sheetId, sheetName, pdf_cnt])
    proc2 = Process(target=t2.start, args=[pdf_cnt])
    proc3 = Process(target=t3.start, args=[sheetId, sheetName, pdf_cnt])
    
    proc1.start()
    proc2.start()
    proc3.start()
    
    proc1.join()
    proc2.join()
    proc3.join()
    
    print("Done!")
    
def run2(sheetId, sheetName):
    sys.stdout = Unbuffered(sys.stdout)

    print("Preparing to run script...")
    service = getGoogleService('gmail', 'v1')
    results = service.users().messages().list(userId='me', labelIds=['INBOX'], q="from:noreply@aqha.org").execute()
    messages = results.get('messages')
    if messages:
        for message in messages:
            msg_id = message['id']
            service.users().messages().delete(userId='me', id=msg_id).execute()

    print("Preparing done.")
    proc1 = Process(target=dt1.start, args=[sheetId, sheetName])
    proc2 = Process(target=dt2.start, args=[sheetId, sheetName])
    
    proc1.start()
    proc2.start()
    
    proc1.join()
    proc2.join()
    
    print("Done!")
    
if __name__ == "__main__":
    while True:
        is_dam = input("Do you want to run this script for only dam name? (y/n): ")
        if is_dam == "y" or is_dam == "n": break
    while True:
        sheetId = input("Enter your worksheet id: ")
        if sheetId != "": break
    while True:
        sheetName = input("Enter your worksheet name: ")
        if sheetName != "": break
        
    if is_dam == "n":
        run(sheetId, sheetName)
    else:
        run2(sheetId, sheetName)