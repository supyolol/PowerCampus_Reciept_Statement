from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import pyodbc
import pandas as pd
from datetime import date
import os
from csv import reader
import time
import win32com.client


##############################################################
##############################################################
## Receipt SQL Logic
##############################################################
##############################################################

## Get Sql Data. Credit Rows. For Statement PDF
def GetCreditRowsReceipt(inputSTUDENTID, inputEntryDate):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select 
            FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as c_entryDate,
            ACADEMIC_YEAR as C_YEAR,ACADEMIC_TERM as C_TERM,
            CRG_CRD_DESC as C_CRG_CRD_DESC,
            CAST(AMOUNT as varchar) as C_AMOUNT 
            from 
            CHARGECREDIT
            where 
            (PEOPLE_ORG_CODE_ID = 'P{studentid}' 
            and ENTRY_DATE = '{entryDate}'
            and CHARGE_CREDIT_TYPE in ('R'))
            or
            (PEOPLE_ORG_CODE_ID in (select ORG_CODE_ID from ORGANIZATION where ORG_NAME_2 = 'P{studentid}') 
            and ENTRY_DATE = '{entryDate}'
            and CHARGE_CREDIT_TYPE in ('R'))

        '''.format(studentid=inputSTUDENTID, entryDate=inputEntryDate)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        test = df.to_dict('records')

        return test

    except Exception as e:
        print(e)


## Get Sql Data. Get Student Info. For Statement PDF.
def GetStudentInfo(inputSTUDENTID):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''select PEOPLE_ID,FIRST_NAME,LAST_NAME,ADDRESS_LINE_1,CITY,STATE,ZIP_CODE from PEOPLE
            left join ADDRESS
            on PEOPLE.PEOPLE_ID = ADDRESS.PEOPLE_ORG_ID
            where PEOPLE_ID = '{studentid}'
        '''.format(studentid=inputSTUDENTID)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        test = df.to_dict('records')

        return test



    except Exception as e:
        print(e)


##############################################################
##############################################################
## Create Receipt and Statment PDF Logic
##############################################################
##############################################################

## Create Receipt
def CreateReceipt(inputSTUDENTID, EntryDateFile):
    # Word Doc Template
    template = "template_receipt_FERPA.docx"

    # Call Fucntion and assign to VAR
    GetCreditRowsVARReceipt = GetCreditRowsReceipt(inputSTUDENTID, EntryDateFile)
    # GetCreditRowsVARReceipt = GetCreditRowsReceipt(inputSTUDENTID,'2021-05-03')
    # Call Fucntion and assign to VAR
    GetStudentInfoVAR = GetStudentInfo(inputSTUDENTID)

    # Remove zeros

    for C in GetCreditRowsVARReceipt:
        amount = float(C['C_AMOUNT'])
        # amount = str(round(amount, 2)) # this was commented out bc need two places after '.'
        amount = "{:.2f}".format(amount)
        C['C_AMOUNT'] = amount

    # Assign Template doc to VAR to do append data to document VAR
    document = MailMerge(template)

    # Append Credit Rows to Word Document
    document.merge_rows('C_CRG_CRD_DESC', GetCreditRowsVARReceipt)

    # student info empty VARs
    F_Name = None
    L_Name = None
    City = None
    State = None
    ZIP = None
    Addy = None
    P_ID = None

    # todays date
    today = date.today().strftime('%m/%d/%Y')
    # for word doc
    string_today = str(today)
    # for file name
    today = date.today().strftime('%m-%d-%Y')
    string_today_file = str(today)
    # student info
    for S in GetStudentInfoVAR:
        F_Name = S['FIRST_NAME']
        L_Name = S['LAST_NAME']
        P_ID = S['PEOPLE_ID']

    # VARs to Doc
    document.merge(

        last_name=L_Name,

        first_name=F_Name,

        TodayDate=string_today,
        peopleID=P_ID,

    )

    # Create Doc
    document.write(str(P_ID) + '_billing_Receipt_' + string_today_file + '.docx')

    # Word Doc to PDF
    wdFormatPDF = 17

    in_file_name = str(P_ID) + '_billing_Receipt_' + string_today_file + '.docx'
    out_file_name = str(P_ID) + '_billing_Receipt_' + string_today_file + '.pdf'

    in_file = os.path.abspath(in_file_name)
    out_file = os.path.abspath(out_file_name)

    word = win32com.client.DispatchEx('Word.Application')
    time.sleep(3)
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    # Delete Word doc
    os.remove(str(P_ID) + '_billing_Receipt_' + string_today_file + '.docx')

    print('CREATED RECEIPT:' + str(P_ID) + '_billing_Receipt_' + string_today_file + '.pdf')


##############################################################
##############################################################
## RUN
##############################################################
##############################################################

# open file in read mode
with open("idscoll.csv", 'r') as read_obj:
    # pass the file object to reader() to get the reader object
    csv_reader = reader(read_obj)
    header = next(csv_reader)
    # Check file as empty
    if header != None:
        # Iterate over each row in the csv using reader object
        for row in csv_reader:
            # row variable is a list that represents a row in csv
            CreateReceipt(row[0], row[3])
