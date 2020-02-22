import click
import dbf
from pathlib import Path
from sqlite_utils import Database
import re
import timeit

docType = 'RA'

def cutString(strIn):
    txt = strIn
    pattern = re.compile(r'\s+')
    txt = re.sub(pattern, '', txt)
    return txt

def getItemList(prNumber):
    path = ".\\temp_db\\OESOIT.DBF"


    table = dbf.Table(str(path))

    try:
        table.open()
    except:
        print('read db error')
    itemList = []

    for row in table:
        dataRow = {
            "index": 0,
            "prNumber": prNumber,
            "itemNumber": "",
            "itemName": "",
            "qty": 0.0,
            "qtCode": "",
            "customer": "",
            "docDate": "",
            "refItem": ""
        }
        if prNumber in str(row[1]) and row[2] != '   ' and row[2] != '@@@':
            dataRow["index"] = int(str(row[2]).replace(' ', ''))
            dataRow["itemNumber"] = str(row[6]).replace('\xa0', ' ').replace('\r\n ', '')
            dataRow["itemName"] = str(row[8]).replace('\xa0', ' ').replace('\r\n ', '')
            dataRow["qty"] = float(cutString(str(row[12])))
            dataRow["qtCode"] = str(row[19])
            dataRow["customer"] = cutString(str(row[5]))
            dataRow["docDate"] = str(row[3])
            dataRow["refItem"] = cutString(str(row[6]).replace('\xa0', ' ').replace('\r\n ', ''))
            itemList.append(dataRow)

    table.close()
    newlist = sorted(itemList, key=lambda k: k["index"])
    return newlist


def getPRList(orderNumber):
    path = ".\\temp_db\\ARTRNRM.DBF"
    docType = 'RA'
    # table_name = Path(path).stem
    table = dbf.Table(str(path))
    table.open()

    prList = []

    for row in table:

        if docType in str(row[0]) and '@2' in str(row[1]) and str(orderNumber) in str(row[2]):
            prNumberBuffer = row[0]
            pattern = re.compile(r'\s+')
            prNumberBuffer = re.sub(pattern, '', prNumberBuffer)
            if not (prNumberBuffer in prList):
                # print(oderBuffer)
                prList.append(prNumberBuffer)
            # print(row2[0])
            # k += 1
        # print(row2[1])
        # if str(row2) == '30 - qty1ny    : 0.0':
        #     print(row2)

    table.close()

    return prList


def getWorkOrderList():
    path = ".\\temp_db\\ARTRNRM.DBF"

    # table_name = Path(path).stem
    table = dbf.Table(str(path))
    table.open()

    orderList = []

    for row in table:

        if docType in str(row[0]) and '@2' in str(row[1]):
            oderBuffer = row[2]
            pattern = re.compile(r'\s+')
            oderBuffer = re.sub(pattern, '', oderBuffer)
            if not (oderBuffer in orderList):
                # print(oderBuffer)
                orderList.append(oderBuffer)
            # print(row2[0])
            # k += 1
        # print(row2[1])
        # if str(row2) == '30 - qty1ny    : 0.0':
        #     print(row2)

    table.close()

    return orderList


def cli1(dbf_paths, sqlite_db, table, verbose):
    """
    Convert DBF files (dBase, FoxPro etc) to SQLite
    https://github.com/simonw/dbf-to-sqlite
    """
    if table and len(dbf_paths) > 1:
        raise click.ClickException("--table only works with a single DBF file")
    # db = Database(sqlite_db)
    # for path in dbf_paths:
    path = dbf_paths
    table_name = table if table else Path(path).stem
    if verbose:
        click.echo('Loading {} into table "{}"'.format(path, table_name))
    table = dbf.Table(str(path))
    table.open()
    columns = table.field_names
    # db[table_name].insert_all(dict(zip(columns, list(row))) for row in table)
    k = 0
    for row2 in table:
        if 'RA' in str(row2[0]) and '@2' in str(row2[1]) and 'B10000239' in str(row2[2]):
            # print(row2[0])
            # print(row2[0])
            k += 1
        # print(row2[1])
        # if str(row2) == '30 - qty1ny    : 0.0':
        #     print(row2)

    table.close()
    # db.vacuum()

# cli1(".\\temp_db\\ARTRNRM.DBF" ,"ko999999999999.db","","")
# for oder in getFromDbf.getWorkOrderList('') :
#     print(oder)


# for item in getFromDbf.getItemList('',k):
#     #print(item["index"],item["itemName"])
#     print(item)

# for pr in getFromDbf.getPRList('', 'B10000290'):
#     print(pr)
#     for item in getFromDbf.getItemList('',pr):
#         print(item)

# sentence = '- - VI0000003'
# sentence.strip()
# sentence.replace(' ', '')
# print(sentence)
