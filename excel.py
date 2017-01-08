#! python3
# Reads an excel file and transforms to JSON.

import xlrd
from collections import OrderedDict
import simplejson as json
import psycopg2
import urllib.parse

#

# set up the connection. Extract the database info from the URI and connect with psycopg2
def connect():
    try:
        result = urllib.parse.urlparse("postgres://4cebab44:@db.nuodata.io:5432/4cebab44")
        username = result.username
        password = result.password
        database = result.path[1:]
        hostname = result.hostname
        DB = psycopg2.connect(
            database = database,
            user = username,
            password = password,
            host = hostname
        )
        c = DB.cursor()
        return DB, c
    except psycopg2.Error as e:
        print(e.pgerror)

# opens the excel file, extracts data row by row, returns it in JSON
def tojson(workbook):
    # Check the file extension is correct
    if workbook.endswith('.xlsx'):
        # Open the workbook.
        wb = xlrd.open_workbook(workbook)
        # select the first sheet.
        sheet1 = wb.sheet_by_index(0)
        # open an empty list where to put the excel data.
        excel_list = []
        # go through each row of excel file adding a dictionry for each
        for rownum in range(sheet1.nrows):
            row_values = sheet1.row_values(rownum)
            products = OrderedDict()
            products['name'] = row_values[0]
            products['age'] = row_values[1]

            excel_list.append(products)
        return excel_list
    else:
        print('The file extension is not supported or is in uppercase.')

# uses psycopg2 to create a query to update the DB.
def post():
    jdata = tojson('practice.xlsx')
    DB, c = connect()
    try:
        for item in jdata:
            c.execute("INSERT INTO test (name, age) VALUES (%s, %s)", (item['name'], item['age']))
    except psycopg2.Error as e:
        DB.rollback()
        print(e.pgerror)
    else:
        DB.commit()
    c.close()
    DB.close()


post()
