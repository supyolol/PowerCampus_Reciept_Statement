from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import pyodbc
import pandas as pd
from datetime import date
from docx2pdf import convert
import os
from csv import reader
import time
import win32com.client


##############################################################
##############################################################
## Statement SQL Logic
##############################################################
##############################################################


## check Credit Records for ANT Flag
def CheckCreditRows(inputSTUDENTID, inputTERM, inputYEAR):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select 
            FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as c_entryDate,
            ACADEMIC_YEAR as C_YEAR,
            ACADEMIC_TERM as C_TERM,
            CRG_CRD_DESC as C_CRG_CRD_DESC,
            CAST(AMOUNT as varchar) as C_AMOUNT,
            ANTICIPATED_FLAG as C_ANT_FLAG
            from 
            CHARGECREDIT
            where 
            PEOPLE_ORG_CODE_ID = 'P{studentid}' 
            and ACADEMIC_TERM = '{term}' 
            and ACADEMIC_YEAR = '{year}'
            and CHARGE_CREDIT_TYPE in ('F','R','D')
        '''.format(studentid=inputSTUDENTID, term=inputTERM, year=inputYEAR)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        data = df.to_dict('records')

        return data

    except Exception as e:
        print(e)


## Get Sql Data. Credit Rows. For Statement PDF
def GetCreditRows(inputSTUDENTID, inputTERM, inputYEAR):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select 
            FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as c_entryDate,
            ACADEMIC_YEAR as C_YEAR,
            ACADEMIC_TERM as C_TERM,
            CRG_CRD_DESC as C_CRG_CRD_DESC,
            CAST(AMOUNT as varchar) as C_AMOUNT,
            ANTICIPATED_FLAG as C_ANT_FLAG
            from 
            CHARGECREDIT
            where 
            PEOPLE_ORG_CODE_ID = 'P{studentid}' 
            and ACADEMIC_TERM = '{term}' 
            and ACADEMIC_YEAR = '{year}'
            and CHARGE_CREDIT_TYPE in ('F','R','D')
            and ANTICIPATED_FLAG <> 'Y'

            ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, term=inputTERM, year=inputYEAR)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        data = df.to_dict('records')

        return data

    except Exception as e:
        print(e)


## Get Sql Data. Debit Rows. For Statement PDF
def GetDebitRows(inputSTUDENTID, inputTERM, inputYEAR):
    servername = '*SERVER*'
    userid = '*USERNAME*'
    password = '*PASSWORD*'
    databasename = '*DATABASE*'

    query = '''
        select 
            FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as d_entryDate,
            ACADEMIC_YEAR as D_YEAR,
            ACADEMIC_TERM as D_TERM,
            CRG_CRD_DESC as D_CRG_CRD_DESC,
            CAST(AMOUNT as varchar) as D_AMOUNT,
            ANTICIPATED_FLAG as D_ANT_FLAG
            from 
            CHARGECREDIT
            where 
            PEOPLE_ORG_CODE_ID = 'P{studentid}' 
            and ACADEMIC_TERM = '{term}' 
            and ACADEMIC_YEAR = '{year}'
            and CHARGE_CREDIT_TYPE in ('C')
            and ANTICIPATED_FLAG <> 'Y'
            ORDER BY ENTRY_DATE ASC


        '''.format(studentid=inputSTUDENTID, term=inputTERM, year=inputYEAR)

    conn = pyodbc.connect(
        'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

    df = pd.read_sql_query(query, conn)

    test = df.to_dict('records')

    return test


## Get Sql Data. Ant Rows. For Statement PDF
def GetANTRows(inputSTUDENTID, inputTERM, inputYEAR):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select 
            FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as a_entryDate,
            ACADEMIC_YEAR as A_YEAR,
            ACADEMIC_TERM as A_TERM,
            CRG_CRD_DESC as A_CRG_CRD_DESC,
            CAST(AMOUNT as varchar) as A_AMOUNT,
            ANTICIPATED_FLAG as A_ANT_FLAG
            from 
            CHARGECREDIT
            where 
            PEOPLE_ORG_CODE_ID = 'P{studentid}' 
            and ACADEMIC_TERM = '{term}'
            and ACADEMIC_YEAR = '{year}'
            and CHARGE_CREDIT_TYPE in ('F','R','D') 
            and ANTICIPATED_FLAG = 'Y'
	    ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, term=inputTERM, year=inputYEAR)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        data = df.to_dict('records')

        return data

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
## Receipt SQL Logic
##############################################################
##############################################################

## Get Sql Data. Credit Rows. For Receipt PDF
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
            PEOPLE_ORG_CODE_ID = 'P{studentid}' 
            and ENTRY_DATE = '{entryDate}'
            and CHARGE_CREDIT_TYPE in ('R')

        '''.format(studentid=inputSTUDENTID, entryDate=inputEntryDate)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        test = df.to_dict('records')

        return test

    except Exception as e:
        print(e)


##############################################################
##############################################################
## Removing Adjusted records for Credits
##############################################################
##############################################################

# Sql data full to only get adjusted records
def GetAdjustedRecords(inputSTUDENTID, inputTERM, inputYEAR):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
        --- only adjustments
select 
FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as c_entryDate,
ACADEMIC_YEAR as C_YEAR,
ACADEMIC_TERM as C_TERM,
CRG_CRD_DESC as C_CRG_CRD_DESC,
ABS(AMOUNT) as C_AMOUNT,
ANTICIPATED_FLAG as C_ANT_FLAG,
CHARGE_CREDIT_CODE
from 
CHARGECREDIT
where 
PEOPLE_ORG_CODE_ID = 'P{studentid}' 
and ACADEMIC_TERM = '{term}' 
and ACADEMIC_YEAR = '{year}'
and CHARGE_CREDIT_TYPE in ('F','R','D')
and ANTICIPATED_FLAG <> 'Y'
and CRG_CRD_DESC like 'Adjusted%'
ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, term=inputTERM, year=inputYEAR)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        test = df.to_dict('records')

        return test

    except Exception as e:
        print(e)


## fucntion for getting key:value data for CC Codes and their Long Descriptions
def GetCODEDes(CODEyo):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select LONG_DESC from CODE_CHARGECREDIT
            where
            CODE_VALUE_KEY = '{x}'
            AND TYPE in ('F','R','D')
            AND STATUS = 'A'
        '''.format(x=CODEyo)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        DataToReturn = df.to_string(index=False, header=False)

        return DataToReturn

    except Exception as e:
        print(e)


## function for getting a single chargecreditnumber id based on multiple paras.
def GetCCpostiveNumber(inputSTUDENTID, ccode, cdesc, term, year, amount):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
Select top 1
CHARGECREDITNUMBER
FROM CHARGECREDIT
where
PEOPLE_ORG_CODE_ID = 'P{studentid}' 
and ACADEMIC_TERM = '{term}' 
and ACADEMIC_YEAR = '{year}'
and CHARGE_CREDIT_TYPE in ('F','R','D')
and ANTICIPATED_FLAG <> 'Y'
and CRG_CRD_DESC = '{cdesc}'
and AMOUNT = '{amount}'
and CHARGE_CREDIT_CODE = '{ccode}'

ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, ccode=ccode, amount=amount, cdesc=cdesc, year=year, term=term)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        DataToReturn = df.to_string(index=False, header=False)

        return DataToReturn

    except Exception as e:
        print(e)


## function to get ChargeCreditNumbers that will be removed from the PDF statement
## for the credits table.
def CCpostiveNumber(inputSTUDENTID, inputTERM, inputYEAR):
    z = GetAdjustedRecords(inputSTUDENTID, inputTERM, inputYEAR)

    for n in z:
        RunCODE = n['CHARGE_CREDIT_CODE']
        LongDescCODE = GetCODEDes(RunCODE)
        if n['CHARGE_CREDIT_CODE'] == RunCODE:
            n['C_CRG_CRD_DESC'] = LongDescCODE

    CCreditNum = []

    for n in z:
        term = n['C_TERM']
        year = n['C_YEAR']
        cdesc = n['C_CRG_CRD_DESC']
        ccode = n['CHARGE_CREDIT_CODE']
        amount = n['C_AMOUNT']
        # lol(ccode,cdesc,term,year,amount):
        x = GetCCpostiveNumber(inputSTUDENTID, ccode, cdesc, term, year, amount)
        CCreditNum.append(x)

    Removeitem = "Empty DataFrame\nColumns: [CHARGECREDITNUMBER]\nIndex: []"
    while Removeitem in CCreditNum:
        CCreditNum.remove("Empty DataFrame\nColumns: [CHARGECREDITNUMBER]\nIndex: []")

    return CCreditNum


##################################################################

## Fucntion to get non adjusted record based on the results of CCPostiveNumber fucntion
## which is the listofNums. Which is a list of chargecreditnumber ids used in the SQL query.
def GetNonAdjustedViaList(inputSTUDENTID, inputTERM, inputYEAR, listofNums):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
select 
FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as c_entryDate,
ACADEMIC_YEAR as C_YEAR,
ACADEMIC_TERM as C_TERM,
CRG_CRD_DESC as C_CRG_CRD_DESC,
('-'+CAST(ABS(AMOUNT) as VARCHAR)) as C_AMOUNT,
ANTICIPATED_FLAG as C_ANT_FLAG,
CHARGE_CREDIT_CODE
from 
CHARGECREDIT
where 
PEOPLE_ORG_CODE_ID = 'P{studentid}' 
and ACADEMIC_TERM = '{term}' 
and ACADEMIC_YEAR = '{year}'
and CHARGE_CREDIT_TYPE in ('F','R','D')
and ANTICIPATED_FLAG <> 'Y'
and CHARGECREDITNUMBER in ({listofNums})
ORDER BY ENTRY_DATE ASC
        '''.format(term=inputTERM, year=inputYEAR, studentid=inputSTUDENTID, listofNums=listofNums)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        test = df.to_dict('records')

        return test

    except Exception as e:
        print(e)


## fucntion for getting key:value data for adjusted CC Codes and their Long Descriptions
def GetCODEDesAdjusted(CODEyo):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select CRG_CRD_DESC
            from CHARGECREDIT
            where CRG_CRD_DESC like 'Adjusted:%'
            and CHARGE_CREDIT_TYPE in ('F','R','D')
            and CHARGE_CREDIT_CODE = '{x2}'
            and ACADEMIC_YEAR > 2018
            GROUP BY CRG_CRD_DESC,CHARGE_CREDIT_CODE
        '''.format(x2=CODEyo)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        DataToReturn = df.to_string(index=False, header=False)

        return DataToReturn

    except Exception as e:
        print(e)


def GetCCnegtiveNumber(inputSTUDENTID, ccode, cdesc, term, year, amount):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
Select
CHARGECREDITNUMBER
FROM CHARGECREDIT
where
PEOPLE_ORG_CODE_ID = 'P{studentid}' 
and ACADEMIC_TERM = '{term}' 
and ACADEMIC_YEAR = '{year}'
and CHARGE_CREDIT_TYPE in ('F','R','D')
and ANTICIPATED_FLAG <> 'Y'
and CRG_CRD_DESC = '{cdesc}'
and AMOUNT = '{amount}'
and CHARGE_CREDIT_CODE = '{ccode}'
ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, ccode=ccode, amount=amount, cdesc=cdesc, year=year, term=term)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        DataToReturn = df.to_string(index=False, header=False)

        return DataToReturn

    except Exception as e:
        print(e)


def CCNegtiveNumber(inputSTUDENTID, inputTERM, inputYEAR, listofNums):
    z = GetNonAdjustedViaList(inputSTUDENTID, inputTERM, inputYEAR, listofNums)

    for n in z:
        RunCODE = n['CHARGE_CREDIT_CODE']
        LongDescCODE = GetCODEDesAdjusted(RunCODE)
        if n['CHARGE_CREDIT_CODE'] == RunCODE:
            n['C_CRG_CRD_DESC'] = LongDescCODE

    CCreditNum = []

    for n in z:
        term = n['C_TERM']
        year = n['C_YEAR']
        cdesc = n['C_CRG_CRD_DESC']
        ccode = n['CHARGE_CREDIT_CODE']
        amount = n['C_AMOUNT']
        # lol(ccode,cdesc,term,year,amount):
        x = GetCCnegtiveNumber(inputSTUDENTID, ccode, cdesc, term, year, amount)
        CCreditNum.append(x)

    Removeitem = "Empty DataFrame\nColumns: [CHARGECREDITNUMBER]\nIndex: []"
    while Removeitem in CCreditNum:
        CCreditNum.remove("Empty DataFrame\nColumns: [CHARGECREDITNUMBER]\nIndex: []")

    return CCreditNum


##################################################################

def CreditRowsWOAdjusted(inputSTUDENTID, inputTERM, inputYEAR, listofNums):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select 
            FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as c_entryDate,
            ACADEMIC_YEAR as C_YEAR,
            ACADEMIC_TERM as C_TERM,
            CRG_CRD_DESC as C_CRG_CRD_DESC,
            CAST(AMOUNT as varchar) as C_AMOUNT,
            ANTICIPATED_FLAG as C_ANT_FLAG
            from 
            CHARGECREDIT
            where 
            PEOPLE_ORG_CODE_ID = 'P{studentid}' 
            and ACADEMIC_TERM = '{term}' 
            and ACADEMIC_YEAR = '{year}'
            and CHARGE_CREDIT_TYPE in ('F','R','D')
            and ANTICIPATED_FLAG <> 'Y'
            and CHARGECREDITNUMBER not in ({listofNums})
            ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, term=inputTERM, year=inputYEAR, listofNums=listofNums)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        data = df.to_dict('records')

        return data

    except Exception as e:
        print(e)


def GetCreditRowsWOAdjusted(inputSTUDENTID, inputTERM, inputYEAR):
    ## Call func to get CC Number (unqiue ids nums)
    q = CCpostiveNumber(inputSTUDENTID, inputTERM, inputYEAR)
    ## Remove "[]" from list var to insert into sql query an 'in'
    qstring = str(q)[1:-1]
    ## User CCpostiveNumber() list of CC Nums to get == adjusted recoreds
    notz = CCNegtiveNumber(inputSTUDENTID, inputTERM, inputYEAR, qstring)

    CCNumbstoremove = q + notz
    CCNumbstoremovestring = str(CCNumbstoremove)[1:-1]

    DATAplzwork = CreditRowsWOAdjusted(inputSTUDENTID, inputTERM, inputYEAR, CCNumbstoremovestring)

    return DATAplzwork


##############################################################
## Removing Adjusted records for Charge
##############################################################
def GetCourseEventID(CourseEventID):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            -- Gets Course EventID from CHARGECREDITCOURSE table.
            select EVENT_ID 
            from CHARGECREDITCOURSE
            where CHARGECREDITNUMBER = '{CourseEventID}'
        '''.format(CourseEventID=CourseEventID)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        
        DataToReturn = df.to_string(index=False, header=False)

        return DataToReturn

    except Exception as e:
        print(e)


# Sql data full to only get adjusted records
def GetAdjustedRecordsCharges(inputSTUDENTID, inputTERM, inputYEAR):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            --- only adjustments
            select 
            CHARGECREDITNUMBER,
            FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as d_entryDate,
            ACADEMIC_YEAR as D_YEAR,
            ACADEMIC_TERM as D_TERM,
            CRG_CRD_DESC as D_CRG_CRD_DESC,
            ABS(AMOUNT) as D_AMOUNT,
            ANTICIPATED_FLAG as D_ANT_FLAG,
            CHARGE_CREDIT_CODE
            from 
            CHARGECREDIT
            where 
            PEOPLE_ORG_CODE_ID = 'P{studentid}' 
            and ACADEMIC_TERM = '{term}' 
            and ACADEMIC_YEAR = '{year}'
            and CHARGE_CREDIT_TYPE in ('C')
            and ANTICIPATED_FLAG <> 'Y'
            and CRG_CRD_DESC like 'Adjusted%'
            ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, term=inputTERM, year=inputYEAR)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        test = df.to_dict('records')

        return test

    except Exception as e:
        print(e)


## fucntion for getting key:value data for CC Codes and their Long Descriptions
def GetLongCODEDesCharges(CODEyo):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select LONG_DESC from CODE_CHARGECREDIT
            where
            CODE_VALUE_KEY = '{x}'
            AND TYPE in ('C')
            AND STATUS = 'A'

        '''.format(x=CODEyo)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        DataToReturn = df.to_string(index=False, header=False)

        return DataToReturn

    except Exception as e:
        print(e)


def GetMedCODEDesCharges(CODEyo):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select MEDIUM_DESC from CODE_CHARGECREDIT
            where
            CODE_VALUE_KEY = '{x}'
            AND TYPE in ('C')
            AND STATUS = 'A'
        '''.format(x=CODEyo)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        DataToReturn = df.to_string(index=False, header=False)

        return DataToReturn

    except Exception as e:
        print(e)


## function for getting a single chargecreditnumber id based on multiple paras.
def GetCCpostiveNumberCharges(inputSTUDENTID, ccode, cdesc, term, year, amount):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
Select top 1
CHARGECREDITNUMBER
FROM CHARGECREDIT
where
PEOPLE_ORG_CODE_ID = 'P{studentid}' 
and ACADEMIC_TERM = '{term}' 
and ACADEMIC_YEAR = '{year}'
and CHARGE_CREDIT_TYPE in ('C')
and ANTICIPATED_FLAG <> 'Y'
--and CRG_CRD_DESC = '{cdesc}'
and AMOUNT = '{amount}'
and CHARGE_CREDIT_CODE = '{ccode}'
ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, ccode=ccode, amount=amount, cdesc=cdesc, year=year, term=term)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        # DataToReturn = df.to_string(index=False,header=False)
        DataToReturn = df['CHARGECREDITNUMBER'].values.tolist()

        return DataToReturn

    except Exception as e:
        print(e)


### for Dup records!
def GetCCpostiveNumberCharges4DUPS(inputSTUDENTID, ccode, cdesc, term, year, amount, DupRecord):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
Select top 1
CHARGECREDITNUMBER
FROM CHARGECREDIT
where
PEOPLE_ORG_CODE_ID = 'P{studentid}' 
and ACADEMIC_TERM = '{term}' 
and ACADEMIC_YEAR = '{year}'
and CHARGE_CREDIT_TYPE in ('C')
and ANTICIPATED_FLAG <> 'Y'
and CRG_CRD_DESC = '{cdesc}'
and AMOUNT = '{amount}'
and CHARGE_CREDIT_CODE = '{ccode}'
and CHARGECREDITNUMBER <> '{DupRecord}'
--and PRINTED_ON_STMT = 'Y'
ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, ccode=ccode, amount=amount, cdesc=cdesc, year=year, term=term,
                   DupRecord=DupRecord)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        # DataToReturn = df.to_string(index=False,header=False)
        DataToReturn = df['CHARGECREDITNUMBER'].values.tolist()

        return DataToReturn

    except Exception as e:
        print(e)


## function to get ChargeCreditNumbers that will be removed from the PDF statement
## for the credits table.
def CCpostiveNumberCharges(inputSTUDENTID, inputTERM, inputYEAR):
    z = GetAdjustedRecordsCharges(inputSTUDENTID, inputTERM, inputYEAR)

    for n in z:

        CCNumVAR = n['CHARGECREDITNUMBER']
        EventID = GetCourseEventID(CCNumVAR)

        if 'Empty DataFrame' not in EventID:
            # print(EventID)
            RunCODE = n['CHARGE_CREDIT_CODE']
            MedDescCODE = GetMedCODEDesCharges(RunCODE)
            if n['CHARGE_CREDIT_CODE'] == RunCODE:
                n['D_CRG_CRD_DESC'] = EventID + ' - ' + MedDescCODE
        else:
            # print("CCNum didin't have a CCCourse Record")
            RunCODE = n['CHARGE_CREDIT_CODE']
            LongDescCODE = GetLongCODEDesCharges(RunCODE)
            if n['CHARGE_CREDIT_CODE'] == RunCODE:
                n['D_CRG_CRD_DESC'] = LongDescCODE

    CCreditNum = []

    for n in z:

        term = n['D_TERM']
        year = n['D_YEAR']
        cdesc = n['D_CRG_CRD_DESC']
        ccode = n['CHARGE_CREDIT_CODE']
        amount = n['D_AMOUNT']

        x = GetCCpostiveNumberCharges(inputSTUDENTID, ccode, cdesc, term, year, amount)
        if x:
            xstring2 = x[0]
        else:
            continue

        if xstring2 in CCreditNum:
            x2 = GetCCpostiveNumberCharges4DUPS(inputSTUDENTID, ccode, cdesc, term, year, amount, xstring2)
            CCreditNum.extend(x2)
        else:
            CCreditNum.extend(x)

    Removeitem = "Empty DataFrame\nColumns: [CHARGECREDITNUMBER]\nIndex: []"
    while Removeitem in CCreditNum:
        CCreditNum.remove("Empty DataFrame\nColumns: [CHARGECREDITNUMBER]\nIndex: []")

    CCreditNum = list(set(CCreditNum))
    # print('PosNums', CCreditNum)
    return CCreditNum


##################################################################

## Fucntion to get non adjusted record based on the results of CCPostiveNumber fucntion
## which is the listofNums. Which is a list of chargecreditnumber ids used in the SQL query.
def GetNonAdjustedViaListCharges(inputSTUDENTID, inputTERM, inputYEAR, listofNums):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select 
            CHARGECREDITNUMBER,
            FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as d_entryDate,
            ACADEMIC_YEAR as D_YEAR,
            ACADEMIC_TERM as D_TERM,
            CRG_CRD_DESC as D_CRG_CRD_DESC,
            ('-'+CAST(ABS(AMOUNT) as VARCHAR)) as D_AMOUNT,
            ANTICIPATED_FLAG as D_ANT_FLAG,
            CHARGE_CREDIT_CODE
            from 
            CHARGECREDIT
            where 
            PEOPLE_ORG_CODE_ID = 'P{studentid}' 
            and ACADEMIC_TERM = '{term}' 
            and ACADEMIC_YEAR = '{year}'
            and CHARGE_CREDIT_TYPE in ('C')
            and ANTICIPATED_FLAG <> 'Y'
            and CHARGECREDITNUMBER in ({listofNums})
            ORDER BY ENTRY_DATE ASC
        '''.format(term=inputTERM, year=inputYEAR, studentid=inputSTUDENTID, listofNums=listofNums)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        test = df.to_dict('records')

        return test

    except Exception as e:
        print(e)


## fucntion for getting key:value data for adjusted CC Codes and their Long Descriptions
def GetCODEDesAdjustedCharges(CODEyo):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select CRG_CRD_DESC
            from CHARGECREDIT
            where CRG_CRD_DESC like 'Adjusted:%'
            and CHARGE_CREDIT_TYPE in ('C')
            and CHARGE_CREDIT_CODE = '{x2}'
            and ACADEMIC_YEAR > 2018
            GROUP BY CRG_CRD_DESC,CHARGE_CREDIT_CODE
        '''.format(x2=CODEyo)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        DataToReturn = df.to_string(index=False, header=False)

        return DataToReturn

    except Exception as e:
        print(e)


def GetCCnegtiveNumberCharges(inputSTUDENTID, ccode, cdesc, term, year, amount):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
Select
CHARGECREDITNUMBER
FROM CHARGECREDIT
where
PEOPLE_ORG_CODE_ID = 'P{studentid}' 
and ACADEMIC_TERM = '{term}' 
and ACADEMIC_YEAR = '{year}'
and CHARGE_CREDIT_TYPE in ('C')
and ANTICIPATED_FLAG <> 'Y'
and CRG_CRD_DESC = '{cdesc}'
and AMOUNT = '{amount}'
and CHARGE_CREDIT_CODE = '{ccode}'
ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, ccode=ccode, amount=amount, cdesc=cdesc, year=year, term=term)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        # DataToReturn = df.to_string(index=False,header=False)
        DataToReturn = df['CHARGECREDITNUMBER'].values.tolist()

        return DataToReturn

    except Exception as e:
        print(e)


def CCNegtiveNumberCharges(inputSTUDENTID, inputTERM, inputYEAR, listofNums):
    z = GetNonAdjustedViaListCharges(inputSTUDENTID, inputTERM, inputYEAR, listofNums)

    for n in z:

        CCNumVAR = n['CHARGECREDITNUMBER']
        # EventID = GetCourseEventID(CCNumVAR)

        # print(EventID)
        RunCODE = n['CHARGE_CREDIT_CODE']
        MedDescCODE = GetMedCODEDesCharges(RunCODE)
        if n['CHARGE_CREDIT_CODE'] == RunCODE:
            # 'Adjusted: Instrumental Music'
            n['D_CRG_CRD_DESC'] = 'Adjusted: ' + MedDescCODE

    CCreditNum = []

    for n in z:
        term = n['D_TERM']
        year = n['D_YEAR']
        cdesc = n['D_CRG_CRD_DESC']
        ccode = n['CHARGE_CREDIT_CODE']
        amount = n['D_AMOUNT']
        # lol(ccode,cdesc,term,year,amount):
        x = GetCCnegtiveNumberCharges(inputSTUDENTID, ccode, cdesc, term, year, amount)
        CCreditNum.extend(x)

    Removeitem = "Empty DataFrame\nColumns: [CHARGECREDITNUMBER]\nIndex: []"
    while Removeitem in CCreditNum:
        CCreditNum.remove("Empty DataFrame\nColumns: [CHARGECREDITNUMBER]\nIndex: []")
    CCreditNum = list(set(CCreditNum))
    # print('NegNums',CCreditNum)
    return CCreditNum


##################################################################

def ChargeRowsWOAdjusted(inputSTUDENTID, inputTERM, inputYEAR, listofNums):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
            select 
            FORMAT(ENTRY_DATE, 'MM/dd/yyyy') as d_entryDate,
            ACADEMIC_YEAR as D_YEAR,
            ACADEMIC_TERM as D_TERM,
            CRG_CRD_DESC as D_CRG_CRD_DESC,
            CAST(AMOUNT as varchar) as D_AMOUNT,
            ANTICIPATED_FLAG as D_ANT_FLAG
            from 
            CHARGECREDIT
            where 
            PEOPLE_ORG_CODE_ID = 'P{studentid}' 
            and ACADEMIC_TERM = '{term}' 
            and ACADEMIC_YEAR = '{year}'
            and CHARGE_CREDIT_TYPE in ('C')
            and ANTICIPATED_FLAG <> 'Y'
            and CHARGECREDITNUMBER not in ({listofNums})
            ORDER BY ENTRY_DATE ASC
        '''.format(studentid=inputSTUDENTID, term=inputTERM, year=inputYEAR, listofNums=listofNums)

        conn = pyodbc.connect(
            'Driver={SQL Server};Server=' + servername + ';UID=' + userid + ';PWD=' + password + ';Database=' + databasename)

        df = pd.read_sql_query(query, conn)

        data = df.to_dict('records')

        return data

    except Exception as e:
        print(e)


def GetChargeRowsWOAdjusted(inputSTUDENTID, inputTERM, inputYEAR):
    ## Call func to get CC Number (unqiue ids nums)
    q = CCpostiveNumberCharges(inputSTUDENTID, inputTERM, inputYEAR)
    ## Remove "[]" from list var to insert into sql query an 'in'
    qstring = str(q)[1:-1]
    ## User CCpostiveNumber() list of CC Nums to get == adjusted recoreds
    notz = CCNegtiveNumberCharges(inputSTUDENTID, inputTERM, inputYEAR, qstring)

    CCNumbstoremove = q + notz

    CCNumbstoremovestring = str(CCNumbstoremove)[1:-1]

    DATAplzwork = ChargeRowsWOAdjusted(inputSTUDENTID, inputTERM, inputYEAR, CCNumbstoremovestring)

    return DATAplzwork


##############################################################
##############################################################
## function to get pervious balnace
##############################################################
##############################################################

def GetCurrentBalance(inputSTUDENTID):
    try:
        servername = '*SERVER*'
        userid = '*USERNAME*'
        password = '*PASSWORD*'
        databasename = '*DATABASE*'

        query = '''
select BALANCE_AMOUNT
from PEOPLEORGBALANCE 
where PEOPLE_ORG_CODE_ID = 'P{studentid}'
and SUMMARY_TYPE = ''
and BALANCE_TYPE = 'CURR'
        '''.format(studentid=inputSTUDENTID)

        conn = pyodbc.connect('Driver={SQL Server};Server='+servername+  ';UID='+userid+';PWD='+password+';Database='+databasename)

        df = pd.read_sql_query(query, conn)

        data = df.to_dict('records')

        return data

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
    template = "template_receipt.docx"

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
        City = S['CITY']
        State = S['STATE']
        ZIP = S['ZIP_CODE']
        Addy = S['ADDRESS_LINE_1']
        P_ID = S['PEOPLE_ID']

    # VARs to Doc
    document.merge(

        street1=Addy,
        last_name=L_Name,
        city=City,
        first_name=F_Name,
        zip=ZIP,
        state=State,
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


## Create Statments
def CreateStatement(inputSTUDENTID, inputTERM, inputYEAR):
    # Ant Flag check
    DaCheck = CheckCreditRows(inputSTUDENTID, inputTERM, inputYEAR)
    if not any(x['C_ANT_FLAG'] == 'Y' for x in DaCheck):

        # Word Doc Template
        template = "template_statement.docx"

        # Call Fucntion and assign to VAR
        GetDebitRowsCHECK = GetDebitRows(inputSTUDENTID, inputTERM, inputYEAR)
        ## check for adjusted
        if any('Adjusted' in x['D_CRG_CRD_DESC'] for x in GetDebitRowsCHECK):

            print('Charge Adjusted Records Found!')

            GetDebitRowsVAR = GetChargeRowsWOAdjusted(inputSTUDENTID, inputTERM, inputYEAR)

        else:
            GetDebitRowsVAR = GetDebitRows(inputSTUDENTID, inputTERM, inputYEAR)
            print('No Charge Adjusted Records Found!')

        # Call Fucntion and assign to VAR, GetCreditRowsVAR
        GetCreditRowsCHECK = GetCreditRows(inputSTUDENTID, inputTERM, inputYEAR)
        ## check for adjusted
        if any('Adjusted' in x['C_CRG_CRD_DESC'] for x in GetCreditRowsCHECK):

            print('Credit Adjusted Records Found!')

            GetCreditRowsVAR = GetCreditRowsWOAdjusted(inputSTUDENTID, inputTERM, inputYEAR)

        else:
            GetCreditRowsVAR = GetCreditRows(inputSTUDENTID, inputTERM, inputYEAR)
            print('No Credit Adjusted Records Found!')

        # Call Fucntion and assign to VAR
        GetStudentInfoVAR = GetStudentInfo(inputSTUDENTID)

        # Remove zeros
        for D in GetDebitRowsVAR:
            amount = float(D['D_AMOUNT'])
            # amount = str(round(amount,2)) # this was commented out bc need two places after '.'
            amount = "{:.2f}".format(amount)
            D['D_AMOUNT'] = amount

        for C in GetCreditRowsVAR:
            amount = float(C['C_AMOUNT'])
            # amount = str(round(amount, 2)) # this was commented out bc need two places after '.'
            amount = "{:.2f}".format(amount)
            C['C_AMOUNT'] = amount

        # Assign Template doc to VAR to do append data to document VAR
        document = MailMerge(template)

        # Append Credit Rows to Word Document
        document.merge_rows('C_CRG_CRD_DESC', GetCreditRowsVAR)
        # Append Debits Rows to Word Document
        document.merge_rows('D_CRG_CRD_DESC', GetDebitRowsVAR)

        # Maths for Credit Total
        sum_total = []
        for T in GetCreditRowsVAR:
            t_int = float(T['C_AMOUNT'])
            sum_total.append(t_int)
        total_to_doc_c = sum(sum_total)
        # string_c_total = str(total_to_doc_c)
        string_c_total = "{:.2f}".format(total_to_doc_c)

        # Maths for Debit Total
        sum_total = []
        for T in GetDebitRowsVAR:
            t_int = float(T['D_AMOUNT'])
            sum_total.append(t_int)
        total_to_doc_d = sum(sum_total)
        # string_d_total = str(total_to_doc_d)
        string_d_total = "{:.2f}".format(total_to_doc_d)

        # Maths for Credits and Debits
        Grand_total = total_to_doc_d - total_to_doc_c

        #Check if Grand total and GetCurrentBalance are equal if not subtract
        CurrentBalance  = GetCurrentBalance(inputSTUDENTID)

        CurrentBalance = CurrentBalance[0]['BALANCE_AMOUNT']

        # string_grand_total = str(round(Grand_total,2))
        string_grand_total = "{:.2f}".format(Grand_total)
        string_CurrentBalance = "{:.2f}".format(CurrentBalance)



        if string_grand_total != string_CurrentBalance:
            print("Does not Eqaul")
            print(f"Grand Total: {string_grand_total}....{type(string_grand_total)}")
            print(f"Current Balance: {string_CurrentBalance}....{type(CurrentBalance)}")

            perivous_balance = float(string_grand_total) - CurrentBalance

            perivous_balance = abs(perivous_balance)

            string_grand_total = float(string_grand_total) + perivous_balance

            string_grand_total = str(string_grand_total)

            perivous_balance = "{:.2f}".format(perivous_balance)

        else:
            perivous_balance = "0.00"







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
            City = S['CITY']
            State = S['STATE']
            ZIP = S['ZIP_CODE']
            Addy = S['ADDRESS_LINE_1']
            P_ID = S['PEOPLE_ID']

        # VARs to Doc
        document.merge(
            C_TERM=inputTERM,
            C_YEAR=inputYEAR,
            street1=Addy,
            last_name=L_Name,
            city=City,
            first_name=F_Name,
            zip=ZIP,
            state=State,
            TodayDate=string_today,
            peopleID=P_ID,
            c_total=string_c_total,
            d_total=string_d_total,
            grand_total=string_grand_total,
            PerBal=perivous_balance

        )

    
        # Create Doc
        document.write(str(P_ID) + '_billing_statment_' + string_today_file + '.docx')

        # Word Doc to PDF
        wdFormatPDF = 17

        in_file_name = str(P_ID) + '_billing_statment_' + string_today_file + '.docx'
        out_file_name = str(P_ID) + '_billing_statment_' + string_today_file + '.pdf'

        in_file = os.path.abspath(in_file_name)
        out_file = os.path.abspath(out_file_name)

        word = win32com.client.DispatchEx('Word.Application')
        time.sleep(1)
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()

        # Delete Word Doc
        os.remove(str(P_ID) + '_billing_statment_' + string_today_file + '.docx')

        print('CREATED STATEMENT:' + str(P_ID) + '_billing_statment_' + string_today_file + '.pdf')

    else:
        # Word Doc Template for ANT FLAG
        template = "template_statement_ANT.docx"

        # Call Fucntion and assign to VAR
        GetDebitRowsCHECK = GetDebitRows(inputSTUDENTID, inputTERM, inputYEAR)
        ## check for adjusted
        if any('Adjusted' in x['D_CRG_CRD_DESC'] for x in GetDebitRowsCHECK):

            print('Charge Adjusted Records Found!')

            GetDebitRowsVAR = GetChargeRowsWOAdjusted(inputSTUDENTID, inputTERM, inputYEAR)

        else:
            GetDebitRowsVAR = GetDebitRows(inputSTUDENTID, inputTERM, inputYEAR)
            print('No Charge Adjusted Records Found!')

        # Call Fucntion and assign to VAR, GetCreditRowsVAR
        GetCreditRowsCHECK = GetCreditRows(inputSTUDENTID, inputTERM, inputYEAR)
        ## check for adjusted
        if any('Adjusted' in x['C_CRG_CRD_DESC'] for x in GetCreditRowsCHECK):

            print('Adjusted Found')

            GetCreditRowsVAR = GetCreditRowsWOAdjusted(inputSTUDENTID, inputTERM, inputYEAR)

        else:
            GetCreditRowsVAR = GetCreditRows(inputSTUDENTID, inputTERM, inputYEAR)
            print('No Adjusted found')

        # Call Function and assign to VAR
        GetANTRowsVAR = GetANTRows(inputSTUDENTID, inputTERM, inputYEAR)
        # Call Fucntion and assign to VAR
        GetStudentInfoVAR = GetStudentInfo(inputSTUDENTID)

        # Remove zeros
        for D in GetDebitRowsVAR:
            amount = float(D['D_AMOUNT'])
            # amount = str(round(amount,2)) # this was commented out bc need two places after '.'
            amount = "{:.2f}".format(amount)
            D['D_AMOUNT'] = amount

        for C in GetCreditRowsVAR:
            amount = float(C['C_AMOUNT'])
            # amount = str(round(amount, 2)) # this was commented out bc need two places after '.'
            amount = "{:.2f}".format(amount)
            C['C_AMOUNT'] = amount

        for A in GetANTRowsVAR:
            amount = float(A['A_AMOUNT'])
            # amount = str(round(amount, 2)) # this was commented out bc need two places after '.'
            amount = "{:.2f}".format(amount)
            A['A_AMOUNT'] = amount

        # Assign Template doc to VAR to do append data to document VAR
        document = MailMerge(template)

        # Append Credit Rows to Word Document
        document.merge_rows('C_CRG_CRD_DESC', GetCreditRowsVAR)
        # Append Debits Rows to Word Document
        document.merge_rows('D_CRG_CRD_DESC', GetDebitRowsVAR)
        # Append Ant Rows to Word Document
        document.merge_rows('A_CRG_CRD_DESC', GetANTRowsVAR)

        # Maths for Credit Total
        sum_total = []
        for T in GetCreditRowsVAR:
            t_int = float(T['C_AMOUNT'])
            sum_total.append(t_int)
        total_to_doc_c = sum(sum_total)
        # string_c_total = str(total_to_doc_c)
        string_c_total = "{:.2f}".format(total_to_doc_c)

        # Maths for Debit Total
        sum_total = []
        for T in GetDebitRowsVAR:
            t_int = float(T['D_AMOUNT'])
            sum_total.append(t_int)
        total_to_doc_d = sum(sum_total)
        # string_d_total = str(total_to_doc_d)
        string_d_total = "{:.2f}".format(total_to_doc_d)

        # Maths for Ant Total
        sum_total = []
        for T in GetANTRowsVAR:
            t_int = float(T['A_AMOUNT'])
            sum_total.append(t_int)
        total_to_doc_a = sum(sum_total)
        # string_d_total = str(total_to_doc_d)
        string_a_total = "{:.2f}".format(total_to_doc_a)

        # Maths for Credits, Debits and Ants
        Grand_total = total_to_doc_d - (total_to_doc_c + total_to_doc_a)

        # Check if Grand total and GetCurrentBalance are equal if not subtract
        CurrentBalance = GetCurrentBalance(inputSTUDENTID)

        CurrentBalance = CurrentBalance[0]['BALANCE_AMOUNT']

        # string_grand_total = str(round(Grand_total,2))
        string_grand_total = "{:.2f}".format(Grand_total)
        string_CurrentBalance = "{:.2f}".format(CurrentBalance)

        if string_grand_total != string_CurrentBalance:

            perivous_balance = float(string_grand_total) - CurrentBalance

            perivous_balance = abs(perivous_balance)

            string_grand_total = float(string_grand_total) + perivous_balance

            string_grand_total = str(string_grand_total)

            perivous_balance = "{:.2f}".format(perivous_balance)

        else:
            perivous_balance = "0.00"

        # string_grand_total = str(round(Grand_total,2))
        #string_grand_total = "{:.2f}".format(Grand_total)

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
            City = S['CITY']
            State = S['STATE']
            ZIP = S['ZIP_CODE']
            Addy = S['ADDRESS_LINE_1']
            P_ID = S['PEOPLE_ID']

        # VARs to Doc
        document.merge(

            C_TERM=inputTERM,
            C_YEAR=inputYEAR,
            street1=Addy,
            last_name=L_Name,
            city=City,
            first_name=F_Name,
            zip=ZIP,
            state=State,
            TodayDate=string_today,
            peopleID=P_ID,
            c_total=string_c_total,
            d_total=string_d_total,
            a_total=string_a_total,
            grand_total=string_grand_total,
            PerBal=perivous_balance

        )


        # Create Doc
        document.write(str(P_ID) + '_billing_statment_' + string_today_file + '.docx')

        # Word Doc to PDF
        wdFormatPDF = 17

        in_file_name = str(P_ID) + '_billing_statment_' + string_today_file + '.docx'
        out_file_name = str(P_ID) + '_billing_statment_' + string_today_file + '.pdf'

        in_file = os.path.abspath(in_file_name)
        out_file = os.path.abspath(out_file_name)

        word = win32com.client.DispatchEx('Word.Application')
        time.sleep(1)
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()

        # Delete Word Doc
        os.remove(str(P_ID) + '_billing_statment_' + string_today_file + '.docx')

        print('CREATED STATEMENT:' + str(P_ID) + '_billing_statment_' + string_today_file + '.pdf')


# *******************************************************************************
# *******************************************************************************
## RUN *************************************************************************
# *******************************************************************************
# *******************************************************************************

# open file in read mode
with open("ids.csv", 'r') as read_obj:
    # pass the file object to reader() to get the reader object
    csv_reader = reader(read_obj)
    header = next(csv_reader)
    # Check file as empty
    if header != None:
        # Iterate over each row in the csv using reader object
        for row in csv_reader:
            # row variable is a list that represents a row in csv
            CreateStatement(row[0], row[1], row[2])
            CreateReceipt(row[0], row[3])
