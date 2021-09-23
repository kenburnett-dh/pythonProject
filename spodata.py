import json
import urllib3
import msal
import pyodbc
import pymssql


def createTables(colMap, valMap):
    # conn = pymssql.connect(
    #     host=r'phx-sql-120\sql2016', database='PowerBI-Test',autocommit=False)
    # cursor = conn.cursor()
    #
    # cursor.execute('drop table if exists SP_Barrow.SP_Enrollments')
    createTblString = 'create table SP_Barrow.SP_Enrollments ' + '(itemid int,'

    invertedColMap = {}
    for col in colMap.items():
        invertedColMap.update({col[1]: col[0]})

    for col in invertedColMap.keys():
        colName = '"' + col.strip() + '"'
        maxlen = 128
        for valDict in valMap.values():
            val = valDict.get(col)
            if not val is None:
                vallen = len(val)
                if vallen > maxlen:
                    maxlen = vallen
        createTblString = createTblString + colName + " varchar(" + str(maxlen) + "),"

    createTblString = createTblString + ")"
    print(createTblString)
    # cursor.execute(createTblString)

    insertValsString = ""
    for id in valMap.items():
        insertHeader = 'insert into SP_Barrow.SP_Enrollments (itemid,"'
        colDict = id[1]
        numvals = len(colDict)
        num = 0
        for col in colDict.keys():
            insertHeader += col
            num += 1
            if num < numvals:
                insertHeader += '","'
        insertHeader += '") values ('
        rowid = id[0]
        print(str(rowid))
        valDict = id[1]
        insertValsString += insertHeader + rowid + ","
        numvals = len(valDict)
        numval = 0;
        for val in valDict.values():
            numval += 1
            val = val.replace("\'", "\'\'")
            insertValsString = insertValsString + "'" + val + "'"
            if numval < numvals:
                insertValsString = insertValsString + ","
        insertValsString = insertValsString + ");"
    insertValsString += "\n"
    # cursor.executemany(insertValsString)
    # cursor.commit()
    # conn.commit()
    # conn.close()


def getToken():
    token = ""
    from msal import ConfidentialClientApplication
    app = ConfidentialClientApplication('a9540a22-0a76-454b-b89b-ee6582bfc045', 'XAy7Q~rHgmCrA5Kdp2pB42tKHbEJT1hGq_00O',
                                        'https://login.microsoftonline.com/078c8807-420a-45d1-85d1-5a0870d06b02/')
    result = app.acquire_token_for_client('https://graph.microsoft.com/.default')
    token = result.get('access_token')
    print(token)
    if token == "":
        token = 'eyJ0eXAiOiJKV1QiLCJub25jZSI6ImhZR0MtTGZsc1pqekMyQWxFYk00T0JRLWczT25hUEp2Ymh6b0dKYm1hWmciLCJhbGciOiJSUzI1NiIsIng1dCI6Imwzc1EtNTBjQ0g0eEJWWkxIVEd3blNSNzY4MCIsImtpZCI6Imwzc1EtNTBjQ0g0eEJWWkxIVEd3blNSNzY4MCJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8wNzhjODgwNy00MjBhLTQ1ZDEtODVkMS01YTA4NzBkMDZiMDIvIiwiaWF0IjoxNjMyMjQ5ODA0LCJuYmYiOjE2MzIyNDk4MDQsImV4cCI6MTYzMjI1MzcwNCwiYWlvIjoiRTJaZ1lEQ3RpK1p6YmZ5dElTU3hZaDNEamJQc0FBPT0iLCJhcHBfZGlzcGxheW5hbWUiOiJPREFUQSIsImFwcGlkIjoiYTk1NDBhMjItMGE3Ni00NTRiLWI4OWItZWU2NTgyYmZjMDQ1IiwiYXBwaWRhY3IiOiIxIiwiaWRwIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvMDc4Yzg4MDctNDIwYS00NWQxLTg1ZDEtNWEwODcwZDA2YjAyLyIsImlkdHlwIjoiYXBwIiwib2lkIjoiNTA1NDRiMTUtZDU3ZC00NDAwLWE4NDYtMjNlYjk2ZjIwNTBmIiwicmgiOiIwLkFXNEFCNGlNQndwQzBVV0YwVm9JY05CckFpSUtWS2wyQ2t0RnVKdnVaWUtfd0VWdUFBQS4iLCJyb2xlcyI6WyJTaXRlcy5SZWFkLkFsbCJdLCJzdWIiOiI1MDU0NGIxNS1kNTdkLTQ0MDAtYTg0Ni0yM2ViOTZmMjA1MGYiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiIwNzhjODgwNy00MjBhLTQ1ZDEtODVkMS01YTA4NzBkMDZiMDIiLCJ1dGkiOiJ6TE5EdnphenUwbWUyWWtIbkxrbkFRIiwidmVyIjoiMS4wIiwid2lkcyI6WyIwOTk3YTFkMC0wZDFkLTRhY2ItYjQwOC1kNWNhNzMxMjFlOTAiXSwieG1zX3RjZHQiOjEzNzgxODA0MDN9.H6Ry_qTu753X2DCJDyDH3oXyZI1haLDsq1US0RFpuON4t9ydisFQHPxqDrpSCUgeCnuhLlJv-LCRtIMZEkl0WwRJ61Zw7u8slzH-KD8hyrozzilWEI5pO6MBDBdwGVL8IrUbNDpo4Nbtxov2lpDE1ONXZeShdQBzc4i3W7028KhJKkdaxkq8mr9B5zizG4FBMxx4B-leFGjOEhQba5rm758vsjzbTyE4IeDBryYE5OeSsebKl-S4x5ZDHZ3VhAnUXI6edyJ4F8-D3FYun3dV8etm0JPtmaKUtI1jFsxmM3WzduuGRAI8bVey_SRjLgb7kIbtJtEdvPYbdK5fj6umfA'
    return token


def getValsJson(json, colMap):
    valMap = {}
    id = ""
    for litems in json:  # dict
        for itemsTuple in litems.items():
            if itemsTuple[0] == 'id':
                id = itemsTuple[1]
            if itemsTuple[0] == 'fields':
                fieldsTuple = itemsTuple[1]
                idValMap = {}
                for field in fieldsTuple.items():  # more tuples
                    fieldName = field[0]
                    fieldVal = field[1]
                    if isinstance(fieldVal, str):
                        col = colMap.get(fieldName)

                        if col is not None:
                            idValMap.update({col: fieldVal})
                valMap.update({id: idValMap})
    return valMap


def getVals(token, colMap):
    http = urllib3.PoolManager()
    valMap = {}
    link = 'https://graph.microsoft.com/v1.0/sites/barrow.sharepoint.com,6c27164b-90ed-47a2-9e0b-e84e9edb2227,f351dd7e-48fc-46c9-8d45-c07ed9d7486b/lists/0c763109-31e1-43ac-a1a0-49936633e709/items?$expand=fields&$top=999'
    response = http.request('GET', link, headers={'Authorization': 'Bearer ' + token})
    odata = json.loads(response.data)
    odataVals = odata.get("value")  # list
    valMap.update(getValsJson(odataVals, colMap))
    link = odata.get("@odata.nextLink")
    while link is not None:
        response = http.request('GET', link, headers={'Authorization': 'Bearer ' + token})
        odata = json.loads(response.data)
        odataVals = odata.get("value")
        valMap.update(getValsJson(odataVals, colMap))
        link = odata.get("@odata.nextLink")
    return valMap


def getCols(token):
    http = urllib3.PoolManager()
    response = http.request('GET',
                            'https://graph.microsoft.com/v1.0/sites/barrow.sharepoint.com,6c27164b-90ed-47a2-9e0b-e84e9edb2227,f351dd7e-48fc-46c9-8d45-c07ed9d7486b/lists/0c763109-31e1-43ac-a1a0-49936633e709/columns',
                            headers={'Authorization': 'Bearer ' + token})
    odata = json.loads(response.data)
    odataValues = odata.get("value")  # list
    colMap = {}  # hold result
    for listval in odataValues:  # dict
        if isinstance(listval, dict):
            internalName = ""
            displayName = ""
            colGroup = ""
            hidden = ""
            for attributeVal in listval.items():  # tuple
                attributeName = attributeVal[0]
                attributeVal = attributeVal[1]
                if attributeName == 'name':
                    internalName = attributeVal
                if attributeName == 'displayName':
                    displayName = attributeVal
                if attributeName == 'columnGroup':
                    colGroup = attributeVal
                if attributeName == 'hidden':
                    hidden = attributeVal
            if colGroup == 'Custom Columns' and not hidden and not internalName.startswith('_'):
                colMap.update({internalName: displayName})
    return colMap


def createTable():
    token = getToken()
    cols = getCols(token)
    vals = getVals(token, cols)
    createTables(cols, vals)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    createTable()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
