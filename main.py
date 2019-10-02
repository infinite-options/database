from flask import (Flask, request, render_template,redirect, url_for,flash, jsonify)
from wtforms import Form, StringField, TextAreaField, PasswordField, validators
import xlrd
import gspread
# will take credential downloaded from google 
from oauth2client.service_account import ServiceAccountCredentials
from xlrd.sheet import ctype_text   
import boto3
import os 
from werkzeug.utils import secure_filename
from flask import send_from_directory
import cx_Oracle
import platform
import requests
import urllib.request
import io
import pandas as pd
from itertools import zip_longest

LOCATION = r"C:\Users\manes\Desktop\DB-InfOpt\instantclient_19_3"
print("ARCH:", platform.architecture())
print("FILES AT LOCATION:")
for name in os.listdir(LOCATION):
    print(name)
os.environ["PATH"] = LOCATION + ";" + os.environ["PATH"]

app = Flask(__name__)

db_address = "infiniteoptions/123456789@dbinfiniteoptions.cxjnrciilyjq.us-west-1.rds.amazonaws.com:1521/ORACLEDB"
connection = cx_Oracle.connect(db_address)

s3 = boto3.client('s3')

@app.route('/')
def index():
    return render_template('home.html')

@app.route('/sheet_name', methods = ['POST'])
def sheet_name():
    bucket = 'bucket-databaseapplication'
    s3 = boto3.client('s3')
    fileName = request.form['excelFileName']
    nameOfSheet = request.form.getlist('client_sheet_name')

    s3.download_file(bucket , fileName, 'C:/Users/manes/Desktop/DB-InfOpt/database/ExcelFile/' + fileName)
    
    firstName = nameOfSheet[0]
    book = xlrd.open_workbook('ExcelFile/' + fileName)
    
    tabName = book.sheet_by_name(firstName)
    
    nameOfColumn = []
    nameOfColumn = tabName.row_values(0)
        
    return render_template('column_name.html', nameOfColumn = nameOfColumn, nameOfSheet = firstName, fileName = fileName)

def getBucket(bucket, fileName):
    s3 = boto3.client('s3')
    obj= s3.get_object(Bucket=bucket, Key=fileName) 
    binary_data = obj['Body'].read()
    book = pd.ExcelFile(io.BytesIO(binary_data))
    sheet = book.sheet_names
    return sheet

@app.route('/upload', methods =['POST'])
def upload():
    
    s3 = boto3.resource('s3')
    
    if bool(request.files.get('fileName', False)) == True:
        file = request.files['myfile']  
        fileName = file.filename
        bucket = 'bucket-databaseapplication'
        s3.Bucket(bucket).put_object(Key=fileName, Body=file)
        sheet = getBucket(bucket, fileName)
    else:
        flash('Please upload a valid file')
        return render_template('home.html')
    
    return render_template('sheetInfo.html', nameOfSheet = sheet, nameOfFile = fileName)

@app.route('/updateData', methods = ['POST'])
def sqlupload():
    
    fileName = request.form['fileName']
    selectedSheet = request.form['selected_sheet']
    
    columnNames = []
    
    columnNames = request.form.getlist('selected_column')
    
    flash("Columns has been uploaded")
    
    #---------------------- refer following code to import data into the database-------------------------#
    
    # book = xlrd.open_workbook('ExcelFile/' + fileName)
    # getBySheetName = book.sheet_by_name(selectedSheet)
    
    # cur = connection.cursor()
    
    # #Drop the table named -- from DB 
    # statement = """    
    #                 BEGIN
    #                     EXECUTE IMMEDIATE 'DROP TABLE ORG_MOCK_INVENTORY4';
    #                 EXCEPTION
    #                     WHEN OTHERS THEN NULL;
    #                 END; 
    #             """
                
    # cur.execute(statement)
    
    # #Create new table 
    # cur.execute(" create table ORG_MOCK_INVENTORY4 (defaultTable VARCHAR(255) NOT NULL)")
    
    # query = "ALTER TABLE ORG_MOCK_INVENTORY4 ADD "
    # dataType = " VARCHAR2(255) NOT NULL "
    
    # for columnName in columnNames:
    #     statement = query + columnName + dataType
    #     print(statement)
    #     cur.execute(statement)
        
    # cur.execute("ALTER TABLE ORG_MOCK_INVENTORY4 DROP UNUSED defaultTable")
    
    # cur.execute(" SELECT gpn,inventory FROM ORG_MOCK_INVENTORY4")

    # # create the sql query called 'insert into' 
    # query = """ INSERT INTO ORG_MOCK_INVENTORY4 (gpn, inventory) VALUES (:1, :2) """    

    # for r in range(1, sheet.nrows):
    
    #     gpn = sheet.cell(r, 0).value
    #     inventory = sheet.cell(r, 1).value
    #     values =(gpn, inventory)
    #     cur.execute(query, values)
    #     #execute sql query 
    
    # cur.execute(""" 
    #                 BEGIN
    #                     EXECUTE IMMEDIATE 'DROP TABLE PM_MOCK_INVENTORY4';
    #                 EXCEPTION
    #                     WHEN OTHERS THEN NULL;
    #                 END; 
    #             """ )

    # cur.execute(""" CREATE TABLE PM_MOCK_INVENTORY4 (
    # gpn VARCHAR(255) NOT NULL,
    # inventory INT NOT NULL
    # ) """)
    
    # #avoid reinserting same element 
    # cur.execute("TRUNCATE TABLE PM_MOCK_INVENTORY4");

    # #sum of same element and insert into new created table 
    # cur.execute(" INSERT INTO PM_MOCK_INVENTORY4 (gpn, inventory) SELECT gpn, SUM(inventory) FROM ORG_MOCK_INVENTORY4 GROUP BY gpn")  

    # # commit the transaction( changes will be permanent)
    # connection.commit()
    
    # #close the connection
    # cur.close()
    # connection.close()
    
    return render_template('sheet_name.html', nameOfSheet = columnNames )

@app.route('/report')
def report():
    # It will return connection object after successfull connection
    # and then it will create a cursor object 
    cur = connection.cursor()

    #this will just returns the number of affected rows otherwise
    # throws an exception
    Resultvalues = cur.execute("SELECT * FROM PM_MOCK_INVENTORY4")

    if Resultvalues > 0:
        labels = [item["gpn"] for item in cur.fetchall()]
    
    Resultlabels = cur.execute("SELECT inventory FROM PM_MOCK_INVENTORY4")

    if Resultlabels > 0:
        values = [item["inventory"] for item in cur.fetchall()] 

    cur.close()

    return render_template('report.html', v = values, l = labels)


@app.route('/oauth')
def oauth():

    #what you wanna modify 
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    credentials = ServiceAccountCredentials.from_json_keyfile_name('g_spreadsheet.json', scope)

    # authorize the client credentials
    gc = gspread.authorize(credentials)

    # open spreadsheet by title first
    wks = gc.open('test')

    #accessing tabs by index 
    gc1 = wks.get_worksheet(0)
    
    dictValorg = gc1.get_all_records()

    #accessing first dictionary
    dictVal = dictValorg[0]

    # make sure you use append method to put all the values into the list
    list = []

    for key in dictVal:
        list.append(key)

    return render_template('oauth.html', list = list)

if __name__=='__main__':
    app.secret_key = 'secret987'
    app.run(host='localhost', port='8080', debug=True)
