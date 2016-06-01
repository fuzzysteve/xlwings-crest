@xw.sub
def facilityIndexs():
    url="https://crest-tq.eveonline.com/industry/systems/"
    wb = Workbook.caller()
    sheet = xw.Sheet('Indexes')
    
    
    r=requests.get(url)
    
    sheetdata=[]
    sheetdata.append(["System Name","System ID","Manufacturing","TE","ME","Copy","Invention"])
    
    if r.status_code == 200:
        for item in r.json()["items"]:
            rowdata=[]
            indexes={}
            for index in item["systemCostIndices"]:
                indexes[index["activityID"]]=index["costIndex"]
            rowdata.append(item["solarSystem"]["name"])
            rowdata.append(item["solarSystem"]["id"])
            rowdata.append(indexes[1])
            rowdata.append(indexes[3])
            rowdata.append(indexes[4])
            rowdata.append(indexes[5])
            rowdata.append(indexes[8])
            sheetdata.append(rowdata)
    Range('A1').value=sheetdata
