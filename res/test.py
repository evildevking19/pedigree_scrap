from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import fitz
import re, time

from constants import *

driver = getGoogleDriver()

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

def findSireFromSite(cn):
    driver.get(f"https://beta.allbreedpedigree.com/search?query_type=check&search_bar=horse&g=5&inbred=Standard&breed=&query={cn.replace(' ', '+')}")
    WebDriverWait(driver, 10).until(lambda browser: browser.execute_script("return document.readyState") == "complete")
    try:
        close_btn = driver.find_element(By.CSS_SELECTOR, "button.btn-close")
        ActionChains(driver).move_to_element(close_btn).click(close_btn).perform()
    except: pass

    time.sleep(0.5)
    try:
        table = driver.find_element(By.CSS_SELECTOR, "table.pedigree-table")
        return getSireNameFromTable(table)
    except:
        try:
            tds = driver.find_elements(By.CSS_SELECTOR, "table.layout-table td[class]:nth-child(1)")
            txt_vals = []
            links = []
            for td in tds:
                txt_vals.append(td.text)
                links.append(td.find_element(By.TAG_NAME, "a").get_attribute("href"))
            indexes = [x for x in txt_vals if x.lower() == cn.lower()]
            print(indexes)
            if len(indexes) == 1:
                driver.get(links[0])
                try:
                    table = driver.find_element(By.CSS_SELECTOR, "table.pedigree-table")
                    return getSireNameFromTable(table)
                except: return ""
            else:
                driver.get(f"https://beta.allbreedpedigree.com/search?match=exact&breed=&sex=&query={cn.replace(' ', '+')}")
                try:
                    table = driver.find_element(By.CSS_SELECTOR, "table.pedigree-table")
                    return getSireNameFromTable(table)
                except: return ""
        except: return ""

def start(file_name):
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
            update_data.append(f"({ext_names[i].lower()})")
        else:
            update_data.append(sire_name)
    print(update_data)
    driver.quit()

start("Order_2023564590.pdf")