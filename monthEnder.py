import sqlalchemy as sql
import os
import win32com.client as win32
import pandas as pd
import pyodbc as sql
import datetime
import warnings

from datetime import date, datetime, timedelta

def getSQLConnectionCursor():

    password = os.getenv('pass')
    username = os.getenv('user')

    global cnxn
    cnxn = sql.connect(DRIVER='{ODBC Driver 17 for SQL Server}', SERVER='172.16.21.75', DATABASE='BI_Finance_Objects', user = username, Password = password)

    global cursor
    cursor = cnxn.cursor()

    return

def monthEnder():

    nowYear = datetime.datetime.now().strftime('%Y')
    nowMonth_Num = datetime.datetime.now().strftime('%M')
    nowMonth_Word = datetime.datetime.now().strftime('%B')



    foldername = "{}\\{}_{}\\".format(nowYear, nowMonth_Num, nowMonth_Word)
    
    path = "P:\\\Shared Services Operations\\\Month End Accrual Reports\\" + foldername
    
    getSQLConnectionCursor()

    getFileInfo = "SELECT * FROM [BI_Finance_Objects].[dbo].[MonthEndAccrualAutomation]"
    
    # format for facility name
    storedProc = "EXEC [BI_Finance_Objects].[dbo].[usp_MonthEndBillingAccrual] '{}', '{}'"


    # Get the complete list of files
    theTrueList = pd.read_sql_query(getFileInfo, cnxn)

    # this is from the True List
    df = pd.read_sql_query(storedProc, cnxn)

    currentime = datetime.datetime.now()

    currentMonth = currentime.strftime('%M')
    currentYear = currentime.strftime('%Y')

    firstDate = get_first_date_of_current_month(currentYear, currentMonth)
    lastDate = get_last_date_of_month(currentYear, currentMonth)
    

    
    for index, item in theTrueList.iterrows():
        
        if item[0] is not None:
            masta = item["Run By"]
            print(masta)

            if ("Facility" in masta ):

                fcltyName = item["Facility"]

                df = pd.read_sql_query(storedProc.format(None,fcltyName), cnxn)

                # filter df by date

                outputTable = df

            if ("Group Name" in masta ):

                grpName = item["Client"]
                df = pd.read_sql_query(storedProc.format(grpName, None), cnxn)

                # filter df by date
                df = df[(df['ShiftDate'] > firstDate) & (df['date'] < lastDate)]
                
                outputTable = df

            if ("Division" in masta ):

                grpName = item["Client"]
                fcltyName = item["Facility"]

                df = pd.read_sql_query(storedProc.format(grpName, None), cnxn)

                df = df[(df['ShiftDate'] > firstDate) & (df['date'] < lastDate)]
                
                # sort by facility name
                newDF = df["fcltyName"]



            today = datetime.date.today()
            first = today.replace(day=1)
            lastMonth = first - datetime.timedelta(days=1)
            lMonth = lastMonth.strftime("%m %Y")

            currentime = datetime.datetime.now()
            mixed = currentime.strftime('%M %Y')
            
            try:
                outputTable = pd.read_sql_query(storedProc.format(item[0]), cnxn)
            except:
                print("Connection Severed")

            outputTable = outputTable.sort_values("Facility")
            outputTable = outputTable.sort_values("CandidateName")
            outputTable = outputTable.sort_values("ShiftDate")
            
            if("Kindred" not in item[0]):
                outputTable.drop('SourceFile', inplace=True, axis=1)

            tableName ="{} Billing Accural {}.xlsx".format(item[0], lMonth)
            
            try:
                writer = pd.ExcelWriter(path + tableName)
                # write dataframe to excel
                outputTable.to_excel(writer, sheet_name='Data', freeze_panes=(1,0), index = False)
                # save the excel
                writer.save()
            except Exception as e:
                print("Connection severed {} ".format(str(e)))

            print('DataFrame is written successfully to Excel File.')
            #outputTable.to_excel(path + tableName)




def sendEmail(subjectLine = None, billToContact = None, billToContact_CC = None, arCollector = None, body = None, filePath = None, chart = None, companyID = None, numberOfInvoices = None):
    #billToContact is a list separated by a semi-colon 
    

    #for testing
    billToContact = "colin.jackson@healthtrustws.com"
    #subjectLine = "TEST Company ID#" xxxxx Test Hospital New Invoice(s) 0 of 0"
    #billToContact_CC = "eric.santovenia@healthtrustws.com; Troy.Raaidy@HealthTrustWS.com"
    billToContact_CC = ""

    


    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = billToContact
        mail.CC = billToContact_CC
        mail.Subject = subjectLine
        mail.HTMLBody = ('Hello All, <br /><br />\nHere is your accrual and billing report.'
        + 'The attached file contains all information.  See column A for:<br /><br />\n'
        + '•	Not Invoiced = accrual amounts<br /><br />\n'
        + '•	Invoiced = billed amounts in the prior month'
        + 'If you have any questions, please do not hesitate to reach out to us.<br /><br />\n'
        + 'Thank you,<br /><br />\n'
        + 'HWS Accounts Receivable HWS.AccountsReceivable@HealthTrustWS.com ')

        #mail.Body = '<h2>HTML Message body</h2>' #this field is optional

        # To attach a file to the email (optional):

        filePath = r"{}".format(filePath)

        #fixes the Albany Bug
        result = os.path.isfile(filePath)
        
        if not result:
            filePath = stripperOmatic(filePath)

        attachment  = filePath
        mail.Attachments.Add(attachment)

        mail.Send()
        Status = "Sent"

    
    except Exception as e:
        #send email to AR team

        Status = ("Not Sent, Critical Email Module Failure. Contact ITG: " + str(e))

    uploadStatusOfSentEmail(filePath, Status)
    
    return

def uploadStatusOfSentEmail():
    print("Sent")


def get_last_date_of_month(year, month):
    """Return the last date of the month.
    
    Args:
        year (int): Year, i.e. 2022
        month (int): Month, i.e. 1 for January

    Returns:
        date (datetime): Last date of the current month
    """
    
    if month == 12:
        last_date = datetime(year, month, 31)
    else:
        last_date = datetime(year, month + 1, 1) + timedelta(days=-1)
    
    return last_date.strftime("%Y-%m-%d")

def get_first_date_of_current_month(year, month):
    """Return the first date of the month.

    Args:
        year (int): Year
        month (int): Month

    Returns:
        date (datetime): First date of the current month
    """
    first_date = datetime(year, month, 1)
    return first_date.strftime("%Y-%m-%d")


monthEnder()