import sqlalchemy as sql
import os
import win32com.client as win32
import pandas as pd
import pyodbc as sql
import datetime
import warnings

def test():
    # format for group name
    storedProc = "EXEC [BI_Finance_Objects].[dbo].[usp_MonthEndBillingAccrual] '{}', null"
    outputTable = pd.read_sql_query(storedProc.format("CHS"), cnxn)


def getSQLConnectionCursor():

    password = os.getenv('pass')
    username = os.getenv('user')

    global cnxn
    cnxn = sql.connect(DRIVER='{ODBC Driver 17 for SQL Server}', SERVER='172.16.21.75', DATABASE='BI_Finance_Objects', user = username, Password = password)

    global cursor
    cursor = cnxn.cursor()

    return
