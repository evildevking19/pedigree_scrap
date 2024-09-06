import fitz
import base64, os, time, sys, re
from constants import getGoogleService, createFileWith

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
   
def getExtactName(org_name):
    org_name = org_name.replace("'", "")
    if re.search(r'\s+\d+', org_name):
        return re.sub(r'\s+\d+', '', org_name).title()
    else:
        return org_name.title()

def extractPdf(file_path):
    NAME_INDEXES = [0, 1, 6, 3, 4, 14, 11, 12, 9, 10, 7, 8, 13, 2]
    NAME_INDEXES2 = [0, 1, 6, 3, 4, 14, 11, 12, 9, 10, 7, 8, 13, 2, 5]
    
    with fitz.open(file_path) as doc:
        rawText = doc[1].get_text()

    rawList = [item for item in rawText.split("\n") if item.strip() != ""]
    ind_start = -1
    try:
        ind_start = rawList.index('Page 1 out of 1')
    except:
        try:
            ind_start = rawList.index('Page 1 out of 2')
        except:
            ind_start = -1
    if ind_start == -1:
        print("Pdf parse error")
        return None
    ind_end = -1
    ind_other = -1
    hasCurrentOwner = True
    for i, text in enumerate(rawList):
        if "CURRENT OWNER" in text:
            ind_end = i
            break
    if ind_end == -1:
        hasCurrentOwner = False
        for i, text in enumerate(rawList):
            if "Horse Details" in text:
                ind_end = i
                ind_other = i
                break
        if ind_end == -1:
            print("Pdf parse error")
            return None
    else:
        for i, text in enumerate(rawList):
            if "Horse Details" in text:
                ind_other = i
                break
            
    data = rawList[ind_start+1:ind_end]
    if ind_end == -1:
        print("Pdf parse error")
        return None
    else:
        tmp_names = []
        tmp_vals = []
        for i, val in enumerate(data):
            if re.search(r'^\d{2}/\d{2}/\d{4}', val):
                if len(tmp_vals) == 3:
                    if i > 3:
                        tmp_names.append(f"{tmp_vals[1]}{tmp_vals[2]}")
                    else:
                        tmp_names.append(f"{tmp_vals[0]}{tmp_vals[1]}{tmp_vals[2]}")
                elif len(tmp_vals) == 2:
                    if i > 2:
                        tmp_names.append(tmp_vals[1])
                    else:
                        tmp_names.append(f"{tmp_vals[0]}{tmp_vals[1]}")
                elif len(tmp_vals) == 1:
                    tmp_names.append(tmp_vals[0])
                tmp_vals = []
            else:
                tmp_vals.append(val)
        names = [None] * 15
        for index, name in enumerate(tmp_names):
            if hasCurrentOwner:
                names[NAME_INDEXES[index]] = getExtactName(name)
            else:
                names[NAME_INDEXES2[index]] = getExtactName(name)
        
        if hasCurrentOwner:
            if len(rawList[ind_other-1].split(" ")) == 1:
                names[5] = getExtactName(rawList[ind_other-3])
            else:
                names[5] = getExtactName(rawList[ind_other-2])
        return names

def checkMailAndDownloadOrderFile(sheetId, sheetName):
    pdf_cnt = 0
    # Create Gmail API service
    serviceForGmail = getGoogleService('gmail', 'v1')
    serviceForSheet = getGoogleService('sheets', 'v4')
    sheetData = serviceForSheet.spreadsheets().values().get(spreadsheetId=sheetId, range=f"{sheetName}!A2:A").execute().get('values')
    while True:
        if os.path.exists("res/dt1.txt"):
            with open("res/dt1.txt", "r") as file:
                c = file.read()
                file.close()
                total_cnt = int(c)
                if pdf_cnt >= total_cnt:
                    os.remove("res/dt1.txt")
                    break
        # Fetch messages from inbox
        results = serviceForGmail.users().messages().list(userId='me', labelIds=['INBOX'], q="from:noreply@aqha.org").execute()
        messages = results.get('messages')
        if not messages:
            print("No messages found.")
        else:
            print("New Messages: " + str(len(messages)))
            for message in messages:
                msg_id = message['id']
                try:
                    msg_body = serviceForGmail.users().messages().get(userId='me', id=msg_id).execute()
                    try:
                        parts = msg_body['payload']['parts']
                        for part in parts:
                            if part['filename']:
                                if 'data' in part['body']:
                                    data = part['body']['data']
                                else:
                                    att_id = part['body']['attachmentId']
                                    att = serviceForGmail.users().messages().attachments().get(userId='me', messageId=msg_id, id=att_id).execute()
                                    data = att['data']
                                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                                filename = part['filename']
                                
                                createFileWith(filename, file_data, 'wb')
                                time.sleep(0.5)
                                ext_names = extractPdf(filename)
                                if ext_names is None: return
                                update_data = []
                                update_data.append([ext_names[2]])
                                
                                for id, row in enumerate(sheetData):
                                    if len(row) != 0:
                                        if ext_names[0].lower() == row[0].lower():
                                            serviceForSheet.spreadsheets().values().update(
                                                spreadsheetId=sheetId,
                                                valueInputOption='RAW',
                                                range=f"{sheetName}!B{str(id+2)}:B{str(id+2)}",
                                                body=dict(
                                                    majorDimension='ROWS',
                                                    values=update_data)
                                            ).execute()
                                os.remove(filename)
                                print("Updated : " + filename)
                                pdf_cnt += 1
                        serviceForGmail.users().messages().delete(userId='me', id=msg_id).execute()
                    except: continue
                except: continue
        time.sleep(5)

def start(sheetId, sheetName):
    sys.stdout = Unbuffered(sys.stdout)
    print("Second process started")
    checkMailAndDownloadOrderFile(sheetId, sheetName)
    print("Second process finished")