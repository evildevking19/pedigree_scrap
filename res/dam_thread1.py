import time, sys, requests
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

def fetchDataFromAQHA(sheetId, sheetName):
    search_cnt = 0
    service = getGoogleService("sheets", "v4")
    horses = service.spreadsheets().values().get(spreadsheetId=sheetId, range=f"{sheetName}!A2:A").execute().get('values')
    service.spreadsheets().values().update(
        spreadsheetId=sheetId,
        valueInputOption='RAW',
        range=f"{sheetName}!A1:B1",
        body=dict(
            majorDimension='ROWS',
            values=[["Horse", "Dam"]])
    ).execute()
    service.spreadsheets().values().clear(
        spreadsheetId=sheetId,
        body={},
        range=f"{sheetName}!B2:B"
    ).execute()
    #### Fetch the all data from website ####
    cnt = 0
    for index, row_data in enumerate(horses):
        if len(row_data) == 0 or row_data[0] == "": continue
        name = row_data[0]
        r = requests.get(f"https://aqhaservices2.aqha.com/api/HorseRegistration/GetHorseRegistrationDetailByHorseName?horseName={name}")
        if r.status_code == 200:
            res = r.json()
            regist_num = res["RegistrationNumber"]
            body_data = {
                # "CustomerEmailAddress": "brittany.holy@gmail.com",
                "CustomerEmailAddress": "pascalmartin973@gmail.com",
                "CustomerID": 1,
                "HorseName": name,
                "RecordOutputTypeCode": "P",
                "RegistrationNumber": regist_num,
                "ReportCode": 21,
                "ReportId": 10008
            }
            r1 = requests.post("https://aqhaservices2.aqha.com/api/FreeRecords/SaveFreeRecord", json=body_data)
            if r1.status_code == 200:
                print("THREAD1: Found Horse (" + name + ") on AQHA Server")
                search_cnt += 1
            else:
                print("THREAD1: Failed to send email from AQHA Server.")
        else:
            print("THREAD1: Not found Horse (" + name + ") on AQHA Server")
        time.sleep(1)
        cnt += 1
        print("Processed " + str(cnt))
    
    createFileWith("res/dt1.txt", str(search_cnt), "w")

def start(sheetId, sheetName):
    sys.stdout = Unbuffered(sys.stdout)
    print("Main process started")
    fetchDataFromAQHA(sheetId, sheetName)
    print("---- Fetch Done ----")
    print("Main process finished")