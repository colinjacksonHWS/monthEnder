import sqlalchemy as sql
import os
import win32com.client as win32
import pandas as pd
import pyodbc as sql
import datetime
import warnings
import re
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

    #pathStart = "P:\\Shared Services Operations\\Month End Accrual Reports\\"
    pathStart = "C:\\Example\\"

    nowYear = datetime.now().strftime('%Y')
    nowMonth_Num = datetime.now().strftime('%m')
    nowMonth_Word = datetime.now().strftime('%B')


    folderName_I = "{}".format(nowYear)
    folderName = "{}\\{}_{}\\".format(nowYear, nowMonth_Num, nowMonth_Word)
    
    #path = "P:\\Shared Services Operations\\Month End Accrual Reports\\" + foldername
    
    path = pathStart + folderName_I
    result = os.path.isdir(path)

    if result is False:
        try:
            os.mkdir(path)
        except Exception as e:
            print("Could not create directory. {}".format(e))

    path = pathStart + folderName


    result = os.path.isdir(path)

    if result is False:
        try:
            os.mkdir(path)
        except Exception as e:
            print("Could not create directory. {}".format(e))
    
    getSQLConnectionCursor()

    getFileInfo = "SELECT * FROM [BI_Finance_Objects].[dbo].[MonthEndAccrualAutomation]"
    
    # format for facility name
    storedProc = "EXEC [BI_Finance_Objects].[dbo].[usp_MonthEndBillingAccrual] '{}', '{}'"


    # Get the complete list of files
    theTrueList = pd.read_sql_query(getFileInfo, cnxn)

    # this is from the True List
    df = pd.read_sql_query(storedProc, cnxn)

    currentime = datetime.now()

    currentMonth = int(currentime.strftime('%m'))
    currentYear = int(currentime.strftime('%Y'))

    firstDate = get_first_date_of_current_month(currentYear, currentMonth)
    lastDate = get_last_date_of_month(currentYear, currentMonth)
    

    
    for index, item in theTrueList.iterrows():
        
        if item[0] is not None:
            masta = item["Run By"]
            print(masta)

            if ("Facility" in masta ):

                fcltyName = item["Facility"]

                try:
                    df = pd.read_sql_query(storedProc.format(None,fcltyName), cnxn)
                except:
                    print("Connection Severed")

                # filter df by date
                df = df[(df['ShiftDate'] > firstDate) & (df['ShiftDate'] < lastDate)]

                outputTable = df
                subject = "{} ~ Accrual and Billing Report".format(fcltyName)

            if ("Group Name" in masta ):

                fcltyName = item["Client"]
                
                try:
                    df = pd.read_sql_query(storedProc.format(grpName, None), cnxn)
                except:
                    print("Connection Severed")

                # filter df by date
                df = df[(df['ShiftDate'] > firstDate) & (df['ShiftDate'] < lastDate)]
                
                outputTable = df
                subject = "{} ~ Accrual and Billing Report".format(fcltyName)

            if ("Division" in masta ):

                grpName = item["Client"]
                fcltyName = item["Facility"]

                try:
                    df = pd.read_sql_query(storedProc.format(grpName, None), cnxn)
                except:
                    print("Connection Severed")

                df = df[(df['ShiftDate'] > firstDate) & (df['ShiftDate'] < lastDate)]
                
                # sort by facility name
                outputTable = df.loc[fcltyName in df["Facility"]]
                subject = "{} ~ Accrual and Billing Report".format(fcltyName)



            today = date.today()
            first = today.replace(day=1)
            lastMonth = first - timedelta(days=1)
            lMonth = lastMonth.strftime("%m %Y")

            thisMonth = first("%m %Y")

            currentime = datetime.now()
            mixed = currentime.strftime('%M %Y')
            

            outputTable = outputTable.sort_values("Facility")
            outputTable = outputTable.sort_values("CandidateName")
            outputTable = outputTable.sort_values("ShiftDate")
            
            if("Kindred" not in item[0]):
                outputTable.drop('SourceFile', inplace=True, axis=1)

            tableName = "{} Billing Accrual {}.xlsx".format(fcltyName, nowMonth_Word)
            
            try:
                pathName = path + tableName
                writer = pd.ExcelWriter(pathName)
                # write dataframe to excel
                outputTable.to_excel(writer, sheet_name='Data', freeze_panes=(1,0), index = False)
                # save the excel
                writer.save()

            except Exception as e:
                print("Connection severed {} ".format(str(e)))

            print('DataFrame is written successfully to Excel File.')
            #outputTable.to_excel(path + tableName)
            
            try:
                email = item["Email"]
                emailCC = item["Email CC"]
            except Exception as e:
                print("Something is wrong with the email")
                continue
            
            if email is None or email is "":
                print("Main Email is empty!")
                print("")
                continue


            sendEmail(subject, email, emailCC, pathName)

def sendEmail(subjectLine = None, billToContact = None, billToContact_CC = None, filePath = None):
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
        mail.HTMLBody = ('Hello, <br /><br />\nHere is your accrual and billing report.'
        + 'The attached file contains all information. See column A for:<br /><br />\n'
        + '•	Not Invoiced = accrual amounts<br /><br />\n'
        + '•	Invoiced = billed amounts in the prior month<br /><br />\n'
        + 'If you have any questions, please do not hesitate to reach out to us.<br /><br />\n'
        + 'Thank you,<br /><br />\n'
        + 'HWS Accounts Receivable HWS.AccountsReceivable@HealthTrustWS.com')

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
        print("Email module failure!")
    
    return

def stripperOmatic(innie = None):

    boom = re.match("^(?:[^\\\\]*\\\\){4}[^\\\\]+", innie)

    newStrang = boom[0].strip()

    outtie = innie.replace(boom[0], newStrang)

    return outtie

def get_last_date_of_month(year, month):
    """Return the last date of the month.
    
    Args:
        year (int): Year, i.e. 2022
        month (int): Month, i.e. 1 for January

    Returns:
        date (datetime): Last date of the current month
    """
    
    if month == 12:
        last_date = date(year, month, 31)
    else:
        last_date = date(year, month + 1, 1) + timedelta(days=-1)
    
    #last_date.strftime("%Y-%m-%d")
    return last_date

def get_first_date_of_current_month(year, month):
    """Return the first date of the month.

    Args:
        year (int): Year
        month (int): Month

    Returns:
        date (datetime): First date of the current month
    """
    first_date = date(year, month, 1)
    #first_date.strftime("%Y-%m-%d")
    return first_date




monthEnder()