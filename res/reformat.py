import sys, time
from constants import getGoogleService, getColumnLabelByIndex

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

def runReformat(psheetId, wsheetId, sheetName):
    service = getGoogleService("sheets", "v4")
    sheetNames = sheetName.split(",")
    for sn in sheetNames:
        print(f"########### Sheet <{sn}> selected ###########")
        wsheetHorses = service.spreadsheets().values().get(spreadsheetId=wsheetId, range=f"{sn}!C2:C").execute().get('values')
        result = service.spreadsheets().values().get(spreadsheetId=psheetId, range=f"{sn}!D8:E").execute().get('values')
        data = ["","","",""]
        for i, r in enumerate(result):
            if i % 3 == 0:
                if data[0] != "":
                    for j, wh in enumerate(wsheetHorses):
                        if wh[0] == data[0]:
                            service.spreadsheets().values().update(
                                spreadsheetId=wsheetId,
                                valueInputOption='RAW',
                                range=f"{sn}!D{j+2}:F{j+2}",
                                body=dict(
                                    majorDimension='ROWS',
                                    values=[[data[1].title(), data[2].title(), data[3].title()]])
                            ).execute()
                            print(f"Updated for <{wh[0]}>.")
                            time.sleep(1)
                            break
                data = ["","","",""]
                data[0] = r[0].rstrip("\xa0")
                sireAndDam = r[1].split(" x ")
                data[1] = sireAndDam[0].rstrip("\xa0")
                data[2] = sireAndDam[1].rstrip("\xa0")
            elif i % 3 == 1:
                data[3] = r[1].rstrip("\xa0").rstrip(")").lstrip("(")

if __name__ == "__main__":
    sys.stdout = Unbuffered(sys.stdout)
    # psheetId = input("Enter producer's sheet id: ")
    # wsheetId = input("Enter your worksheet id: ")
    # sheetName = input("Enter multiple sheet names as you need (seperate by comma): ")
    # runReformat(psheetId, wsheetId, sheetName)
    runReformat("1NaLxb9TRwMFeIC9zSOsCDVYcovSuGQ-oRMFIl9mm6aI", "1RZEoNp-vRG5Y0D7btXBuHpzDHtRtIVPI58P8jO89tt8", "Fut 1st Go")
    print("---- Done ----")