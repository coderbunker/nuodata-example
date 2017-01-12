#! python3
# Reads an excel file and transforms to JSON.

import xlrd
import sys
from collections import OrderedDict
import simplejson as json
import psycopg2
import argparse

# TODO update readme with what this does, and usage example.
# empty comment on line 10
# TODO fix custom exception line 29
# useless undesirable check
# TODO make parameter for database function and for post function.

# set up the connection. Extract the database info from the URI and connect with psycopg2
def takesArguments():
    # parses the args containing the excel file directory and stores it into a variable.
    parser = argparse.ArgumentParser(description='Sends information from an excel sheet to the nuodata DB.')
    parser.add_argument('excel-file', help='the directory of the excel file you want to send to the db.')
    parser.add_argument('database uri', help='this includes the uri of the db. include all required parameters in the uri.')

    excel = parser.parse_args()
    # Asks user for database api information.
    database = input('database name please: ')
    username = input('username please: ')
    password = input('password please: ')
    hostname = input('hostname please: ')

def connect():
    # connect with these credentials.
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
        return('Could not connect plese check that your credentials have been entered correctly')

# opens the excel file, extracts data row by row, returns it in JSON
def tojson(workbook):
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
