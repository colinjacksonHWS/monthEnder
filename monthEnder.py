import sqlalchemy as sql
import os
import win32com.client as win32
import pandas as pd
import pyodbc as sql
import datetime
import warnings


warnings.filterwarnings('ignore',
 r"^Dialect sqlite\+pysqlite does \*not\* support Decimal objects natively\, "
 "and SQLAlchemy must convert from floating point - rounding errors and other "
 "issues may occur\. Please consider storing Decimal numbers as strings or "
 "integers on this platform for lossless storage\.$")



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
    
    getExamples = "SELECT DISTINCT dcm.GroupName FROM BI_DataMart..DimCompanyMaster dcm WHERE dcm.Business = 'External' ORDER BY 1"
    storedProc = "EXEC [BI_Finance_Objects].[dbo].[usp_MonthEndBillingAccrual] '{}'"

    df = pd.read_sql_query(getExamples, cnxn)

    for index, item in df.iterrows():
        if item[0] is not None:

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
        mail.HTMLBody = ('Dear Valued Client, <br /><br />\nThank you for doing business with HealthTrust Workforce Solutions.'
        + ' Attached please find your new invoice(s).<br />' 
        + "<br />{}".format(chart)
        + "<br /><b><i>Did you know...</b></i> we accept payments via Check, ACH, or EFT? "
        + "For any question or concerns regarding the attached items, please reach out to: HWS.AccountsReceivable@HealthTrustWS.com"
        + "<br /><br />\n\nSincerely,<br />"
        + arCollector 
        + ", Finance Shared Services<br />HWS.AccountsReceivable@HealthTrustWS.com<br />HealthTrust Workforce Solutions | 1000 Sawgrass Corp Pkwy, 6th Floor | Sunrise, FL 33323"
        + "<br />Click <a href=\"http://engage.healthtrustjobs.com/rate-your-healthtrust-experience\">here</a> to rate your HealthTrust experience!")

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


monthEnder()