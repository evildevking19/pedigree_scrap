import fitz
import re

def getExtactName(org_name):
    org_name = org_name.replace("'", "")
    if re.search(r'\s+\d+', org_name):
        return re.sub(r'\s+\d+', '', org_name).title()
    else:
        return org_name.title()

def extractPdf(file_path):
    with fitz.open(file_path) as doc:
        rawText = doc[0].get_text()

    print(rawText)
    rawList = [item for item in rawText.split("\n") if item.strip() != ""]

ext_names = extractPdf("2.pdf")