import re, fitz, os
from constants import *

LETTER_COLORS = ["Bay", "Bay Overo", "Bay Roan", "Black", "Black Overo", "Black Roan", "Black/Brown", "Blue Roan", "Brown", "Buckskin", "Buckskin Roan", "Champagne", "Chestnut", "Chestnut Overo", "Chestnut Roan", "Cremello", "Dark Bay/Brown", "Dark Chestnut", "Dun", "Dunalino", "Dunskin", "Gray", "Gray/Roan", "Grulla", "Liver Chestnut", "Overo", "Palomino", "Perlino", "Pinto", "Red Dun", "Red Roan", "Roan", "Sabino", "Silver Dapple", "Smoky Grulla", "Sorrel", "White"]

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
    

files = os.listdir("res/orders")
for file in files:
    # print("################################")
    # print(file)
    print(extractPdf(f"res/orders/{file}"))
    # print("################################")