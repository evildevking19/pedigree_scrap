import pycurl
from bs4 import BeautifulSoup
from io import BytesIO
from constants import *

headers = ["User-Agent: Python-PycURL", "Accept: */*"]

def searchNameFromABP(horse_name):
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
        print(getSheetDataFrom(table))
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
                    print(getSheetDataFrom(table))
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
                    print(getSheetDataFrom(table))
                else: print("THREAD1: Not found (" + horse_name + ") in https://beta.allbreedpedigree.com/")
        else: print("THREAD1: Not found (" + horse_name + ") in https://beta.allbreedpedigree.com/")

searchNameFromABP("Frenchman’s Dashing High")