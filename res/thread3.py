import fitz
import re, time, sys
import pycurl
from io import BytesIO
from bs4 import BeautifulSoup

from constants import *

LETTER_COLORS = ["Bay", "Bay Overo", "Bay Roan", "Black", "Black Overo", "Black Roan", "Black/Brown", "Blue Roan", "Brown", "Buckskin", "Buckskin Roan", "Champagne", "Chestnut", "Chestnut Overo", "Chestnut Roan", "Cremello", "Dark Bay/Brown", "Dark Chestnut", "Dun", "Dunalino", "Dunskin", "Gray", "Gray/Roan", "Grulla", "Liver Chestnut", "Overo", "Palomino", "Perlino", "Pinto", "Red Dun", "Red Roan", "Roan", "Sabino", "Silver Dapple", "Smoky Grulla", "Sorrel", "White"]

worksheet = None
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
        tmp_val = ""
        for i, val in enumerate(data):
            if val.title() in LETTER_COLORS: continue
            if re.search(r'^\d{2}/\d{2}/\d{4}', val):
                tmp_names.append(tmp_val)
                tmp_val = ""
            else:
                tmp_val += val
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

def findSireFromSite(cn):
    horse_name = cn.replace(" ", "+").replace("'", "").encode("ascii", "ignore").decode("utf-8")
    buffer = BytesIO()
    c = pycurl.Curl()
    c.setopt(c.URL, f'https://beta.allbreedpedigree.com/search?query_type=check&search_bar=horse&g=5&inbred=Standard&breed=&query={horse_name}')
    c.setopt(c.HTTPHEADER, headers)
    c.setopt(pycurl.FOLLOWLOCATION, 1)
    c.setopt(c.WRITEDATA, buffer)
    c.perform()
    c.close()
    r = buffer.getvalue().decode('utf-8')
    soup = BeautifulSoup(r, 'html.parser')
    table = soup.select_one("table.pedigree-table")
    if table != None:
        return getSireNameFromTable(table)
    else:
        tds = soup.select("table.layout-table td[class]:nth-child(1)")
        if tds != None or len(tds) != 0:
            txt_vals = []
            links = []
            for td in tds:
                txt_vals.append(td.text)
                links.append(td.select_one("a").get("href"))
            indexes = [x.strip() for x in txt_vals if x.strip().lower() == cn.lower()]
            if len(indexes) == 1:
                buffer = BytesIO()
                c = pycurl.Curl()
                c.setopt(c.URL, links[0])
                c.setopt(pycurl.FOLLOWLOCATION, 1)
                c.setopt(c.HTTPHEADER, headers)
                c.setopt(c.WRITEDATA, buffer)
                c.perform()
                c.close()
                r1 = buffer.getvalue().decode('utf-8')
                soup1 = BeautifulSoup(r1, 'html.parser')
                table = soup1.select_one("table.pedigree-table")
                if table != None:
                    return getSireNameFromTable(table)
                else: return ""
            else:
                buffer = BytesIO()
                c = pycurl.Curl()
                c.setopt(c.URL, f'https://beta.allbreedpedigree.com/search?match=exact&breed=&sex=&query={horse_name}')
                c.setopt(c.HTTPHEADER, headers)
                c.setopt(pycurl.FOLLOWLOCATION, 1)
                c.setopt(c.WRITEDATA, buffer)
                c.perform()
                c.close()
                r2 = buffer.getvalue().decode('utf-8')
                soup2 = BeautifulSoup(r2, 'html.parser')
                table = soup2.select_one("table.pedigree-table")
                if table != None:
                    return getSireNameFromTable(table)
                else: return ""
        else: return ""

def updateGSData(file_name, sheetId, sheetName, indexOfHorse, sheetData, multichoices):
    file_path = ORDER_DIR_NAME + "/" + file_name
    update_data = []
    ext_names = extractPdf(file_path)
    if ext_names is None: return

    pre_index_list = [1, 2, 5, 3, 5, 7, 11, 9, 13]
    for i in pre_index_list:
        update_data.append(ext_names[i])
    
    for i in range(7, 15):
        sire_name = findSireFromSite(ext_names[i])
        if sire_name.strip() == "":
            tmp_name = ""
            for choice_val in multichoices:
                if choice_val[0] == f"({ext_names[i].lower()})":
                    if len(choice_val) > 1:
                        tmp_name = choice_val[1]
            if tmp_name != "":
                update_data.append(tmp_name)
            else:
                update_data.append(f"({ext_names[i].lower()})")
        else:
            update_data.append(sire_name)

    for id, row in enumerate(sheetData):
        if len(row) != 0:
            if ext_names[0].lower() == row[indexOfHorse].lower():
                worksheet.values().update(
                    spreadsheetId=sheetId,
                    valueInputOption='RAW',
                    range=f"{sheetName}!{getColumnLabelByIndex(indexOfHorse+1)}{str(id+2)}:Z{str(id+2)}",
                    body=dict(
                        majorDimension='ROWS',
                        values=[update_data])
                ).execute()
                time.sleep(1)
    print("THREAD3: " + file_name)
    os.remove(file_path)

def start(sheetId, sheetName, init_cnt):
    sys.stdout = Unbuffered(sys.stdout)
    print("Third process started")
    createOrderDirIfDoesNotExists()
    global worksheet
    service = getGoogleService("sheets", "v4")
    worksheet = service.spreadsheets()
    multichoices = worksheet.values().get(spreadsheetId=sheetId, range=f"mult choices!A2:B").execute().get('values')
    # try:
    file_cnt = init_cnt
    while True:
        if os.path.exists("res/t2.txt"):
            if os.path.exists("res/t1.txt"):
                with open("res/t1.txt", "r") as file:
                    c = file.read()
                    file.close()
                    total_cnt = int(c)
                    if file_cnt >= total_cnt:
                        os.remove("res/t1.txt")
                        os.remove("res/t2.txt")
                        break
        files = getOrderFiles()
        if len(files) > 0:
            values = worksheet.values().get(spreadsheetId=sheetId, range=f"{sheetName}!A1:Z").execute().get('values')
            header = values.pop(0)
            indexOfHorseHeader = header.index('Horse')
            for file in files:
                updateGSData(file, sheetId, sheetName, indexOfHorseHeader, values, multichoices)
                file_cnt += 1

    parenthesis = []
    values = worksheet.values().get(spreadsheetId=sheetId, range=f"{sheetName}!A1:Z").execute().get('values')
    for row in values:
        if len(row) != 0:
            for col_val in row:
                if re.match(r'\(.*?\)', col_val):
                    filtered_choices = [x for x in multichoices if x[0] == col_val]
                    if len(filtered_choices) == 0:
                        if len(parenthesis) == 0:
                            parenthesis.append([col_val, ""])
                        else:
                            filtered_value = [x for x in parenthesis if x[0] == col_val]
                            if len(filtered_value) == 0:
                                parenthesis.append([col_val, ""])
    
    worksheet.values().update(
        spreadsheetId=sheetId,
        valueInputOption='RAW',
        range=f"mult choices!A{str(len(multichoices)+2)}:B",
        body=dict(
            majorDimension='ROWS',
            values=parenthesis)
    ).execute()
    # except:
    #     print("There is no <mult choices> sheet!")

# start("13b-fBnZpZFC_PTTuJ0Y9pYA-UYIgbsUDCCHjga5RBzs", "horses", 0)