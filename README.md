## Synopsis
This file takes an excel file and sends it to the specified URI database. In this case it's made for sending updates to a database on nuodata.io.

## Pre-reqs
Run the following in the cloned directory to install the required packages:
```
pip install -r requirements.txt
```

## Running
```
python3 excel.py $EXCEL_FILE $DATABASE_URI $DATABASE_NAME
python3 excel.py -h
python3 excel.py --help
```
Be sure that the the URI contains all the information needed to access the Database
For example:
```
postgres://[username:password@]host[:PORT]/username
```

## Motivation
This was made as practice to a bigger project that sends the same type of data to an API.

## Installation
git clone repo

## Reference
- tojson method found on: http://www.anthonydebarros.com/2013/02/05/get-json-from-excel-using-python-xlrd/

- error handling seen on: http://stackoverflow.com/questions/8497886/graceful-primary-key-error-handling-in-python-psycopg2
