#! python3
# Reads an excel file and transforms to JSON.

import xlrd
import sys
import simplejson as json
from collections import OrderedDict
import psycopg2
import argparse
import urllib.parse

# TODO update readme with what this does, and usage example.
# empty comment on line 10
# TODO fix custom exception line 29
# useless undesirable check
# TODO make parameter for database function and for post function.

# set up the connection. Extract the database info from the URI and connect with psycopg2

################################################################################
def main():
    parser = argparse.ArgumentParser(description='Sends information from an excel sheet to the nuodata DB.')
    parser.add_argument('xcel', help='The directory of the excel file you want to send to the db.')
    parser.add_argument('uri', help='This includes the uri of the db. include all required parameters in the uri.')
    parser.add_argument('dbname', help='the name of the database')

    args = parser.parse_args()

    xcel = args.xcel
    uri = args.uri
    dbname = args.dbname

    # connect to the database using connect() (see below)
    DB, c= connect(uri)

    # get the list of rows from the excel file
    excel_list = toJson(xcel)

    # send the items in the list to the db
    post(excel_list)

    # close the connection to the cursor and then to the database
    c.close()
    DB.close()
##############################################
# THIS PART CONNECTS TO THE DB AND CREATES A cursor
# connect with these credentials.
def connect(uri):
    # Asks user for database api information.
    result = urllib.parse.urlparse(uri)
    database = result.username
    username = result.password
    password = result.path[1:]
    hostname = result.hostname
    try:
        DB = psycopg2.connect(
            database = database,
            user = username,
            password = password,
            host = hostname
        )
        c = DB.cursor()
        return DB, c
    except:
         print('Could not connect please check that your credentials have been entered correctly.')

##############################################
# THIS PART opens the excel file, extracts data row by row, returns it as a list
# Open the workbook.
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

##############################################
# uses psycopg2 to create a query to update the DB.
def post(excel_list):
    try:
        for item in jdata:
            c.execute("INSERT INTO %s (product_name, quantity) VALUES (%s, %s)", (database-name, item['name'], item['age']))
    except psycopg2.Error as e:
        DB.rollback()
        print(e.pgerror)
    else:
        DB.commit()

if __name__ == '__main__':
    main()
