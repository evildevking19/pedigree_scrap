import time, sys, requests
import pycurl
from io import BytesIO
from bs4 import BeautifulSoup
from constants import *

headers = ["User-Agent: Python-PycURL", "Accept: */*"]

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

def searchNameFromABP(service, sheetId, sheetName, indexOfHorseHeader, horse_name, index):
    sheet_data = []
    horse_name = horse_name.replace(" ", "+").replace("'", "").encode("ascii", "ignore").decode("utf-8")
    buffer = BytesIO()
    c = pycurl.Curl()
    c.setopt(c.URL, f'https://beta.allbreedpedigree.com/search?query_type=check&search_bar=horse&g=5&inbred=Standard&breed=&query={horse_name}')
    c.setopt(c.HTTPHEADER, headers)
    c.setopt(c.WRITEDATA, buffer)
    c.perform()
    c.close()
    r = buffer.getvalue().decode('utf-8')
    soup = BeautifulSoup(r, 'html.parser')
    table = soup.select_one("table.pedigree-table")
    if table != None:
        sheet_data.append(getSheetDataFrom(table))
    else:
        tds = soup.select("table.layout-table td[class]:nth-child(1)")
        if tds != None or len(tds) != 0:
            txt_vals = []
            links = []
            for td in tds:
                txt_vals.append(td.text)
                links.append(td.select_one("a").get("href"))
            indexes = [x.strip() for x in txt_vals if x.strip().lower() == horse_name.lower()]
            if len(indexes) == 1:
                buffer = BytesIO()
                c = pycurl.Curl()
                c.setopt(c.URL, links[0])
                c.setopt(c.HTTPHEADER, headers)
                c.setopt(c.WRITEDATA, buffer)
                c.perform()
                c.close()
                r1 = buffer.getvalue().decode('utf-8')
                soup1 = BeautifulSoup(r1, 'html.parser')
                table = soup1.select_one("table.pedigree-table")
                if table != None:
                    sheet_data.append(getSheetDataFrom(table))
                else: print("THREAD1: Not found (" + horse_name + ") in https://beta.allbreedpedigree.com/")
            else:
                buffer = BytesIO()
                c = pycurl.Curl()
                c.setopt(c.URL, f'https://beta.allbreedpedigree.com/search?match=exact&breed=&sex=&query={horse_name}')
                c.setopt(c.HTTPHEADER, headers)
                c.setopt(c.WRITEDATA, buffer)
                c.perform()
                c.close()
                r2 = buffer.getvalue().decode('utf-8')
                soup2 = BeautifulSoup(r2, 'html.parser')
                table = soup2.select_one("table.pedigree-table")
                if table != None:
                    sheet_data.append(getSheetDataFrom(table))
                else: print("THREAD1: Not found (" + horse_name + ") in https://beta.allbreedpedigree.com/")
        else: print("THREAD1: Not found (" + horse_name + ") in https://beta.allbreedpedigree.com/")

    if len(sheet_data) != 0:
        service.spreadsheets().values().update(
            spreadsheetId=sheetId,
            valueInputOption='RAW',
            range=f"{sheetName}!{getColumnLabelByIndex(indexOfHorseHeader+1)}{str(index+2)}:Z{str(index+2)}",
            body=dict(
                majorDimension='ROWS',
                values=sheet_data)
        ).execute()

def fetchDataFromAQHA(sheetId, sheetName, init_cnt):
    search_cnt = init_cnt
    global browser
    service = getGoogleService("sheets", "v4")
    result = service.spreadsheets().values().get(spreadsheetId=sheetId, range=f"{sheetName}!A1:Z").execute().get('values')
    header = result.pop(0)
    indexOfHorseHeader = header.index("Horse")
    service.spreadsheets().values().clear(
        spreadsheetId=sheetId,
        body={},
        range=f"{sheetName}!{getColumnLabelByIndex(indexOfHorseHeader+1)}1:Z"
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=sheetId,
        valueInputOption='RAW',
        range=f"{sheetName}!{getColumnLabelByIndex(indexOfHorseHeader+1)}1:Z1",
        body=dict(
            majorDimension='ROWS',
            values=[["Sire","Dam","Dams Sire","Grandsire Top","Grandsire Bottom","Grandsire Sire Top","Grandsire Sire Bottom","Granddams Sire Top","Granddams Sire bottom","Great-Grandsires Sire Top (1)","Great-Granddams Sire Top (2)","Great-Grandsires Sire Top (3)","Great-Granddams Sire Top (4)","Great-Grandsires Sire Bottom (5)","Great-Granddams Sire Bottom (6)","Great-Grandsires Sire Bottom (7)","Great-Granddams Sire Bottom (8)"]])
    ).execute()
    #### Fetch the all data from website ####
    cnt = 0
    for index, row_data in enumerate(result):
        if len(row_data) == 0 or row_data[indexOfHorseHeader] == "": continue
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
            searchNameFromABP(service, sheetId, sheetName, indexOfHorseHeader, name, index)
        time.sleep(1)
        cnt += 1
        print("Processed " + str(cnt))
    
    createFileWith("res/t1.txt", str(search_cnt), "w")

def start(sheetId, sheetName, init_cnt):
    sys.stdout = Unbuffered(sys.stdout)
    print("Main process started")
    fetchDataFromAQHA(sheetId, sheetName, init_cnt)
    print("---- Fetch Done ----")
    print("Main process finished")