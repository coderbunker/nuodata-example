#! python3
# Reads an excel file and sends the row information to a table using  postresql.

import xlrd
import sys
import simplejson as json
from collections import OrderedDict
import psycopg2
import argparse
import urllib.parse


def main():
    parser = argparse.ArgumentParser(description='Sends information from an excel sheet to the nuodata DB.')
    parser.add_argument('xcel', help='The directory of the excel file you want to send to the db.')
    parser.add_argument('uri', help='This includes the uri of the db. include all required parameters in the uri.')
    parser.add_argument('tableName', help='the name of the database')

    args = parser.parse_args()

    xcel = args.xcel
    uri = args.uri
    tableName = args.tableName

    # connect to the database using connect() (see below)
    DB, c = connect(uri)

    # get the list of rows from the excel file
    excel_list = toJson(xcel)

    # send the items in the list to the db
    post(DB, c, tableName, excel_list)

    # close the connection to the cursor and then to the database
    c.close()
    DB.close()


# connect with these credentials.
def connect(uri):
    try:
        DB = psycopg2.connect(uri)
        c = DB.cursor()
        return DB, c
    except:
         raise Exception('Could not connect please check that your credentials have been entered correctly.', sys.exc_info()[0])


# Opens the excel file, extracts data row by row, returns it as a list
def toJson(xcel):
    wb = xlrd.open_workbook(xcel)
    # select the first sheet.
    sheet1 = wb.sheet_by_index(0)
    # open an empty list where to put the excel data.
    excel_list = []
    # go through each row of excel file adding a dictionry for each
    for rownum in range(sheet1.nrows):
        row_values = sheet1.row_values(rownum)
        products = OrderedDict()
        products['product_name'] = row_values[0]
        products['quantity'] = row_values[1]

        excel_list.append(products)
    return excel_list


# uses psycopg2 to create a query to update the DB.
def post(DB, c, tableName, excel_list):
    try:
        for item in excel_list:
            query = "INSERT INTO %s (product_name, quantity) VALUES ('%s', %s)" % (tableName, item['product_name'], item['quantity'])
            print (query)
            c.execute(query)
    except psycopg2.Error as e:
        DB.rollback()
        print(e.pgerror)
    else:
        DB.commit()

if __name__ == '__main__':
    main()
