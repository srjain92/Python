import pyodbc
import pandas as pd
import xlsxwriter
import time
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from xlsxwriter import worksheet

today = datetime.date.today()
first = today.replace(day=1)
lastMonth = first - datetime.timedelta(days=1)
lastMonth = (lastMonth.strftime("%B"))

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=dds-sql03;'
                      'Database=DLPLUS;'
                      'Trusted_Connection=yes;')

script1 = ('''WITH CTE1
            AS
            (SELECT  MONTH(C.[DateInvoiced]) AS [Month]
                   ,'DDS Lab' AS [Lab]
                   ,COALESCE(A17.[CustomerCategory1],P.[MajorGroupID],'Other') AS [Product Class]
                   ,COALESCE(A17.[CustomerCategory2],P.[MajorGroupID],'Other') AS [Product Class 2]
                   ,CASE WHEN COALESCE(PH.[Analysis17],'') = '' THEN CP.[ProductID] ELSE COALESCE(PH.[Analysis17],'') END AS [Product ID]
                   ,COALESCE(A17.[CustomerDescription],P.[Description],'') AS [Product Name]
                   ,CAST(SUM(CP.[Quantity]) AS FLOAT) AS [Quantity]
                   ,CAST(SUM(COALESCE(CP.[TotalCharge],0) - COALESCE(CP.[TotalTax],0)) AS FLOAT) AS [Total Spend]
                   ,CASE WHEN DI.[CaseID] IS NOT NULL 
                         THEN 'Digital'
                         ELSE 'Traditional'
                    END AS [Case Type]
            
            FROM [DLPLUS].[dbo].[Cases] C (NOLOCK)
                   LEFT JOIN [DLPLUS].[dbo].[CaseProducts] CP (NOLOCK) ON CP.[CaseID] = C.[CaseID]
                   LEFT JOIN [DLPLUS].[dbo].[Products] P (NOLOCK) ON P.[ProductID] = CP.[ProductID]
                   LEFT JOIN [DLPLUS].[dbo].[Doctors] D (NOLOCK) ON D.[DoctorID] = C.[DoctorID]
                   LEFT JOIN [LabTrac].[dbo].[Product_Header] PH (NOLOCK) ON PH.[AltDescription] = CP.[ProductID]
                   LEFT JOIN [LabTrac].[dbo].[DDS_Analysis17Mappings] A17 (NOLOCK) ON A17.[CustomerCode] = PH.[Analysis17]
                   LEFT JOIN (SELECT
                                    C.[CaseID]
                                FROM [DLPlus].[dbo].[Cases] C (NOLOCK)
                                    LEFT JOIN [DLPLUS].[dbo].[CaseMaterials] CM (NOLOCK) ON CM.[CaseID] = C.[CaseID]
                                    LEFT JOIN [DLPlus].[dbo].[CaseProducts] CP (NOLOCK) ON CP.[CaseID] = C.[CaseID]
                                    LEFT JOIN [DLPlus].[dbo].[Products] P (NOLOCK) ON P.[ProductID] = CP.[ProductID]
                                WHERE
                                    C.[TYPE] = 0
                                    AND C.[DateIn] IS NOT NULL
                                    AND (CM.[MaterialID] = 'DIGITAL_IMPRESSION'
                                    OR CP.[ProductID] IN ('DIGITAL_PRINT_MODEL','ITERO_MODELS','MODELESS','DITERO','DMODPRINT','DSIRONA','DTRIOS','DOOP','DDDX','DCSCONNECT','DCHECKDIE'))
                                GROUP BY
                                    C.[CaseID]) DI ON DI.[CaseID] = C.[CaseID]
            WHERE
                   YEAR(C.[DateInvoiced]) = YEAR(DATEADD(M,-1,GETDATE()))
                   AND MONTH(C.[DateInvoiced]) = MONTH(DATEADD(M,-1,GETDATE()))
                   AND D.[BillTo] = 'HEARTLAND_DE32014775'
                   AND (CASE WHEN COALESCE(PH.[Analysis17],'') = '' THEN CP.[ProductID] ELSE COALESCE(PH.[Analysis17],'') END) NOT IN ('CHARGE_REMAKE', 'CHARGE_REMAKE_NOREA','CHARGE_REPAIR','COMBOCASE',
                        'DIMPDIGPRT','IMODEVAL','IMODIFYABU','IMPL_CASE','IMPLREMAKE','INTPROCEED','IPKGANALOG','IREPAIRCOM','NO_CHARGE_REMAKE','NOCALLOOP','NOCALLPL','','PROCEED_AS_IS','REMOVABLE_CASE',
                        'RREMCHARGE','Yes','ISCRWDESC','D3MPORTAL','DCSCONNECT','DITERO','DMODPRINT','DSIRONA','DTRIOS','ERA LABOR','F34CROWN','FIXREMAKE','ICHANNELE','IFINSMETLI','IFINSMETOC','IDIETRIM',
                        'MODELESS','NEW_IMPRESSION','NO_CHARGE_REPAIR','ORTHO_CASE','OUT_WARRANTY','RELINE_UNDER_WARR','REP_UNDER_WARRAN_C&B','REPAIR_UNDER_WARR','RESET_RETRY_1','RESET_RETRY_2','RREMAKEREP')
            
            GROUP BY
                   CASE WHEN COALESCE(PH.[Analysis17],'') = '' THEN CP.[ProductID] ELSE COALESCE(PH.[Analysis17],'') END
                   ,COALESCE(A17.[CustomerDescription],P.[Description],'')
                   ,COALESCE(A17.[CustomerCategory1],P.[MajorGroupID],'Other')
                   ,MONTH(C.[DateInvoiced])
                   ,COALESCE(A17.[CustomerCategory2],P.[MajorGroupID],'Other')
                   ,CASE WHEN DI.[CaseID] IS NOT NULL 
                         THEN 'Digital'
                         ELSE 'Traditional'
                    END
            )
            
            SELECT [Month]
                  ,[Lab]
                  ,[Product Class]
                  ,[Product Class 2]
                  ,[Product ID]
                  ,[Product Name]
                  ,COALESCE([Digital Quantity],0) AS [Digital Quantity]
                  ,COALESCE([Digital Total Spend],0) AS [Digital Total Spend]
                  ,COALESCE([Traditional Quantity],0) AS [Traditional Quantity]
                  ,COALESCE([Traditional Total Spend],0) AS [Traditional Total Spend]
            
            FROM (SELECT [Month]
                              ,[Lab]
                              ,[Product Class]
                              ,[Product Class 2]
                              ,[Product ID]
                              ,[Product Name]
                              ,[Case Type] + ' ' + [Type] AS [Type]
                              ,VALUE
                        FROM CTE1
            
                        UNPIVOT
                        (
                            VALUE
                            FOR [Type] in ([Quantity],[Total Spend])
                        ) AS UnPivotExample) UP
            
            PIVOT 
            (
             SUM(VALUE)
             FOR [Type] IN ([Digital Quantity], [Digital Total Spend],[Traditional Quantity], [Traditional Total Spend])
            ) AS PIVOTTABLE
            
            ORDER BY [Product ID]''')

script2 = ('''DECLARE @StartDate DATE = DATEADD(month, DATEDIFF(month, 0, getdate()) - 1, 0)
              DECLARE @EndDate DATE = DATEADD(month, DATEDIFF(month, 0, getdate()) , 0) 

                SELECT
                 'DDS Lab' AS [Lab Name]
                 ,COALESCE(D.[PracticeName],'') AS [Office Name]
                 ,C.[PatientFirst] + ' ' + C.[PatientLast] AS [Customer Name]
                 ,replace(replace(ID.[Data], char(10), ''), char(13), '') AS [Epicor ID]
                 ,CASE WHEN COALESCE(PH.[Analysis17],'') = '' THEN CP.[ProductID] ELSE COALESCE(PH.[Analysis17],'') END AS [Product ID]
                 ,COALESCE(A17.[CustomerDescription],P.[Description],'') AS [Product Name]
                 ,SUM(CASE WHEN MODELLESS.[CaseID] IS NOT NULL AND ([dbo].[Weekday_Date_Diff_ExcludingHolidays](C.[DateIn], COALESCE(FedExIPDScans.[ShipDate],CD.[ShipToCustomerDate],C.[DateInvoiced])) + 1) <= 3 
                           THEN CP.Quantity 
                           WHEN (IMPLANTPRODUCT.[CaseID] IS NOT NULL OR IMPLANTMATERIAL.[CaseID] IS NOT NULL) AND ([dbo].[Weekday_Date_Diff_ExcludingHolidays](C.[DateIn], COALESCE(FedExIPDScans.[ShipDate],CD.[ShipToCustomerDate],C.[DateInvoiced])) + 1) <= 11 
                           THEN CP.Quantity 
                           WHEN ([dbo].[Weekday_Date_Diff_ExcludingHolidays](C.[DateIn], COALESCE(FedExIPDScans.[ShipDate],CD.[ShipToCustomerDate],C.[DateInvoiced])) + 1) <= 8 
                           THEN CP.Quantity 
                           ELSE 0 END) AS [Units Delivered within TAT]
                 ,COALESCE(CAST(SUM(CASE WHEN MODELLESS.[CaseID] IS NOT NULL AND ([dbo].[Weekday_Date_Diff_ExcludingHolidays](C.[DateIn], COALESCE(FedExIPDScans.[ShipDate],CD.[ShipToCustomerDate],C.[DateInvoiced])) + 1) <= 3 
                           THEN CP.Quantity 
                           WHEN (IMPLANTPRODUCT.[CaseID] IS NOT NULL OR IMPLANTMATERIAL.[CaseID] IS NOT NULL) AND ([dbo].[Weekday_Date_Diff_ExcludingHolidays](C.[DateIn], COALESCE(FedExIPDScans.[ShipDate],CD.[ShipToCustomerDate],C.[DateInvoiced])) + 1) <= 11 
                           THEN CP.Quantity 
                           WHEN ([dbo].[Weekday_Date_Diff_ExcludingHolidays](C.[DateIn], COALESCE(FedExIPDScans.[ShipDate],CD.[ShipToCustomerDate],C.[DateInvoiced])) + 1) <= 8 
                           THEN CP.Quantity 
                           ELSE 0 END) AS FLOAT)/CAST(NULLIF(SUM(CP.Quantity),0) AS FLOAT) * 100,0) AS [On-Time Delivery%]
                 ,SUM(CP.Quantity) AS [Total Units Made]
                 ,COALESCE(SUM(CASE WHEN C.RemakeOf IS NOT NULL THEN CP.Quantity ELSE NULL END),0) AS [Remake Units]
                 ,COALESCE(CAST(COALESCE(SUM(CASE WHEN C.RemakeOf IS NOT NULL THEN CP.Quantity ELSE NULL END),0) AS FLOAT)/CAST(NULLIF(SUM(CP.Quantity),0) AS FLOAT) * 100,0) AS [Remake Unit%]
                 ,COALESCE(SUM(CASE WHEN CP.UNITPRICE <> 0 THEN CP.Quantity ELSE NULL END),0)  AS [Units Billed]
                 ,SUM(CP.UNITPRICE*CP.Quantity) AS [AmountBilled]
                
                FROM Cases C (NOLOCK) 
                 INNER JOIN Doctors D (NOLOCK) ON C.DoctorID = D.DoctorID 
                 LEFT JOIN [DLPLUS].[dbo].[UserDefinedFieldData] ID (NOLOCK) ON ID.[KeyFieldData] = D.DoctorID AND ID.[UserDefinedFieldID] = 1
                 INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID 
                 INNER JOIN Products P (NOLOCK) ON CP.[ProductID] = P.[ProductID] 
                 LEFT JOIN [LabTrac].[dbo].[Product_Header] PH (NOLOCK) ON PH.[AltDescription] = CP.[ProductID]
                 LEFT JOIN [LabTrac].[dbo].[DDS_Analysis17Mappings] A17 (NOLOCK) ON A17.[CustomerCode] = PH.[Analysis17]
                 LEFT JOIN [DLPLUS].[dbo].[CaseDates] CD (NOLOCK) ON CD.[CaseID] = C.[CaseID]
                 LEFT JOIN (SELECT 
                                    FDX.[CaseID]
                                    ,MAX(FDX.[ID]) AS [LastResultID]
                              FROM [DDSLab_Services].[dbo].[International_FedExScans] FDX (NOLOCK)
                                    INNER JOIN [DLPLUS].[dbo].[CASES] C (NOLOCK) ON C.[CaseID] = FDX.[CaseID] 
                              WHERE
                                    FDX.[Deleted] = 0
                              GROUP BY
                                    FDX.[CaseID]) AS LastFedExIPDScan ON LastFedExIPDScan.[CaseID] = C.[CaseID]
                 LEFT JOIN [DDSLab_Services].[dbo].[International_FedExScans] FedExIPDScans (NOLOCK) ON FedExIPDScans.[Id] = LastFedExIPDScan.[LastResultID]
                 LEFT JOIN (SELECT
                                        C.[CaseID]
                                    FROM [DLPlus].[dbo].[Cases] C (NOLOCK)
                                        LEFT JOIN [DLPlus].[dbo].[CaseProducts] CP (NOLOCK) ON CP.[CaseID] = C.[CaseID]
                                        LEFT JOIN [DLPlus].[dbo].[Products] P (NOLOCK) ON P.[ProductID] = CP.[ProductID]
                                    WHERE
                                        C.[TYPE] = 0
                                        AND C.[DateIn] IS NOT NULL
                                        AND CP.[ProductID] IN ('MODELESS', 'DCHECKDIE')
                                    GROUP BY
                                        C.[CaseID]) MODELLESS ON MODELLESS.[CaseID] = C.[CaseID]
                LEFT JOIN (SELECT
                                        C.[CaseID]
                                    FROM [DLPlus].[dbo].[Cases] C (NOLOCK)
                                        LEFT JOIN [DLPlus].[dbo].[CaseMaterials] CM (NOLOCK) ON CM.[CaseID] = C.[CaseID]
                                    WHERE
                                        C.[TYPE] = 0
                                        AND C.[DateIn] IS NOT NULL
                                        AND (CM.[MaterialID] = 'IMP_PARTS_ENCLOSED' OR CM.[MaterialID] = 'IMPLANT_RESTORATION')
                                    GROUP BY
                                        C.[CaseID]) IMPLANTMATERIAL ON IMPLANTMATERIAL.[CaseID] = C.[CaseID]
                LEFT JOIN (SELECT
                                        C.[CaseID]
                                    FROM [DLPlus].[dbo].[Cases] C (NOLOCK)
                                        LEFT JOIN [DLPlus].[dbo].[CaseProducts] CP (NOLOCK) ON CP.[CaseID] = C.[CaseID]
                                        LEFT JOIN [DLPlus].[dbo].[Products] P (NOLOCK) ON P.[ProductID] = CP.[ProductID]
                                    WHERE
                                        C.[TYPE] = 0
                                        AND C.[DateIn] IS NOT NULL
                                        AND (P.[UnitGroupID] = 'IMPL' OR P.[ProductID] LIKE 'IMPL_%')
                                    GROUP BY
                                        C.[CaseID]) IMPLANTPRODUCT ON IMPLANTPRODUCT.[CaseID] = C.[CaseID]
                
                WHERE D.BillTo = 'HEARTLAND_DE32014775'
                 AND C.DateInvoiced >= @StartDate
                 AND C.DateInvoiced < @EndDate
                 AND C.[TYPE] = 0
                 AND CP.[ProductID] NOT IN  ('CHARGE_REMAKE', 'CHARGE_REMAKE_NOREA','CHARGE_REPAIR','COMBOCASE',
                'DIMPDIGPRT','IMODEVAL','IMODIFYABU','IMPL_CASE','IMPLREMAKE','INTPROCEED','IPKGANALOG','IREPAIRCOM','NO_CHARGE_REMAKE','NOCALLOOP','NOCALLPL','','PROCEED_AS_IS','REMOVABLE_CASE',
                'RREMCHARGE','Yes','ISCRWDESC','D3MPORTAL','DCSCONNECT','DITERO','DMODPRINT','DSIRONA','DTRIOS','ERA LABOR','F34CROWN','FIXREMAKE','ICHANNELE','IFINSMETLI','IFINSMETOC','IDIETRIM',
                'MODELESS','NEW_IMPRESSION','NO_CHARGE_REPAIR','ORTHO_CASE','OUT_WARRANTY','RELINE_UNDER_WARR','REP_UNDER_WARRAN_C&B','REPAIR_UNDER_WARR','RESET_RETRY_1','RESET_RETRY_2','RREMAKEREP',
                'SHIPPING','RETURN_PRODUCTSHIP')
                AND PH.[Analysis17] <> 'YES'
                 AND P.[MajorGroupID] <> 'SHIPPING' 
                
                GROUP BY COALESCE(D.[PracticeName],'') 
                 ,C.[PatientFirst] + ' ' + C.[PatientLast] 
                 ,CASE WHEN COALESCE(PH.[Analysis17],'') = '' THEN CP.[ProductID] ELSE COALESCE(PH.[Analysis17],'') END 
                 ,COALESCE(A17.[CustomerDescription],P.[Description],'')
                 ,replace(replace(ID.[Data], char(10), ''), char(13), '') 
                
                ORDER BY [Product ID]''')

script3 = ('''DECLARE @StartDate DATE = DATEADD(month, DATEDIFF(month, 0, getdate()) - 1, 0)
              DECLARE @EndDate DATE = DATEADD(month, DATEDIFF(month, 0, getdate()) , 0) 

                SELECT 
                       'DDSLab' AS [LabName]
                       ,replace(replace(ID.[Data], char(10), ''), char(13), '') AS [Epicor ID]
                       ,D.PracticeName AS [Office Name]
                       ,C.[PatientFirst] + ' ' + C.[PatientLast] AS [Customer Name]
                       ,C.CaseID AS [Invoice Number]
                       ,CASE WHEN COALESCE(PH.[Analysis17],'') = '' THEN CP.[ProductID] ELSE COALESCE(PH.[Analysis17],'') END AS [Product ID]
                       ,COALESCE(A17.[CustomerDescription],P.[Description],'') AS [Product Name]
                       ,CP.QUANTITY AS [Remake Units]
                       ,CASE WHEN (C.TOTALCHARGE - C.TOTALTAX) < 10 
                             THEN 'No Charge'  
                             WHEN (C.TOTALCHARGE - C.TOTALTAX) >= 10 
                             THEN 'Charge'
                        END AS [Charge/No Charge]
                       ,COALESCE(REPLACE(REPLACE(REASON.Description, char(10), ''), char(13), ''), RTRIM(RR.RemakeLevel1) + ' - ' + RTRIM(RR.RemakeLevel2) + ' ' + RTRIM(RR.RemakeLevel3),R1.Reason, 'Other') As [Remake Reason] 
                
                FROM   CASES C  (NOLOCK)
                       INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID
                       LEFT JOIN [DLPLUS].[dbo].[UserDefinedFieldData] ID (NOLOCK) ON ID.[KeyFieldData] = D.DoctorID AND ID.[UserDefinedFieldID] = 1
                       LEFT JOIN CASEPRODUCTS CP (NOLOCK) ON CP.CASEID = C.CASEID
                       LEFT JOIN PRODUCTS P (NOLOCK) ON P.PRODUCTID = CP.PRODUCTID  
                       LEFT JOIN DDS_QCImpressionScore S (nolock) on S.[CaseID] = C.[CaseID] 
                       LEFT JOIN DDS_QCRemakeScore RS (nolock) on RS.[CaseID] = C.[CaseID]          
                       LEFT JOIN DDS_QCRemakeItemDetail RD (nolock) on RD.QCRemakeScoreID = RS.QCRemakeScoreID 
                       LEFT JOIN DDS_QCRemakeReasons RR (nolock) on RR.ID = RD.QCRemakeItemID 
                       LEFT JOIN Reasons R1 (nolock) on C.[ReasonID] = R1.[ReasonID] 
                       LEFT JOIN [LabTrac].[dbo].[DDS_Order_Stage_DLPLUSLINK] OSD (NOLOCK) ON OSD.CASEID = C.CASEID
                       LEFT JOIN [LabTrac].[dbo].[Order_Stage_Products] OSP (NOLOCK) ON OSP.[OrderID] = OSD.[OrderID] AND OSP.[StageID] = OSD.[StageID] AND OSP.REMAKE = 1
                       LEFT JOIN [LabTrac].[dbo].[Order_Stage_Product_Remakes] OSPR (NOLOCK) ON OSPR.[StageItemID] = OSP.[Ref]
                       LEFT JOIN [LabTrac].[dbo].[Category_Defintions] REASON (NOLOCK) ON REASON.Ref = OSPR.ReasonID
                       LEFT JOIN [LabTrac].[dbo].[Product_Header] PH (NOLOCK) ON PH.[AltDescription] = CP.[ProductID]
                       LEFT JOIN [LabTrac].[dbo].[DDS_Analysis17Mappings] A17 (NOLOCK) ON A17.[CustomerCode] = PH.[Analysis17]
                
                WHERE D.BillTo = 'HEARTLAND_DE32014775'
                AND C.REMAKEOF IS NOT NULL
                AND C.DateInvoiced >= @StartDate
                AND C.DateInvoiced < @EndDate
                AND C.[TYPE] = 0
                AND CP.[ProductID] NOT IN  ('CHARGE_REMAKE', 'CHARGE_REMAKE_NOREA','CHARGE_REPAIR','COMBOCASE',
                'DIMPDIGPRT','IMODEVAL','IMODIFYABU','IMPL_CASE','IMPLREMAKE','INTPROCEED','IPKGANALOG','IREPAIRCOM','NO_CHARGE_REMAKE','NOCALLOOP','NOCALLPL','','PROCEED_AS_IS','REMOVABLE_CASE',
                'RREMCHARGE','ISCRWDESC','D3MPORTAL','DCSCONNECT','DITERO','DMODPRINT','DSIRONA','DTRIOS','ERA LABOR','F34CROWN','FIXREMAKE','ICHANNELE','IFINSMETLI','IFINSMETOC','IDIETRIM',
                'MODELESS','NEW_IMPRESSION','NO_CHARGE_REPAIR','ORTHO_CASE','OUT_WARRANTY','RELINE_UNDER_WARR','REP_UNDER_WARRAN_C&B','REPAIR_UNDER_WARR','RESET_RETRY_1','RESET_RETRY_2','RREMAKEREP')
                AND PH.[Analysis17] <> 'YES'
                AND P.[MajorGroupID] <> 'SHIPPING' 
                
                GROUP BY
                       D.PracticeName 
                       ,replace(replace(ID.[Data], char(10), ''), char(13), '') 
                       ,C.[PatientFirst] + ' ' + C.[PatientLast] 
                       ,C.CaseID 
                       ,CASE WHEN COALESCE(PH.[Analysis17],'') = '' THEN CP.[ProductID] ELSE COALESCE(PH.[Analysis17],'') END 
                       ,COALESCE(A17.[CustomerDescription],P.[Description],'') 
                       ,CP.QUANTITY 
                       ,COALESCE(REPLACE(REPLACE(REASON.Description, char(10), ''), char(13), ''), RTRIM(RR.RemakeLevel1) + ' - ' + RTRIM(RR.RemakeLevel2) + ' ' + RTRIM(RR.RemakeLevel3),R1.Reason, 'Other')
                       ,CASE WHEN (C.TOTALCHARGE - C.TOTALTAX) < 10 
                             THEN 'No Charge'    
                             WHEN (C.TOTALCHARGE - C.TOTALTAX) >= 10 
                             THEN 'Charge'
                        END
                
                ORDER BY [Product ID]''')

### Create a Pandas dataframe from the data
df1 = pd.read_sql_query(script1, conn)
df2 = pd.read_sql_query(script2, conn)
df3 = pd.read_sql_query(script3, conn)

filename = 'Heartland_Monthly_Reporting' + '_' + lastMonth + '.xlsx'
sheetname1 = 'Quantity Report'
sheetname2 = 'Delivery and Quality Report'
sheetname3 = 'Remake Report'

### Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(filename, engine='xlsxwriter')

### Convert the dataframe to an XlsxWriter Excel object
df1.to_excel(writer, sheet_name=sheetname1, index=False)
df2.to_excel(writer, sheet_name=sheetname2, index=False)
df3.to_excel(writer, sheet_name=sheetname3, index=False)

# create function for excel column auto width
def auto_width_column(df):
    for i, col in enumerate(df.columns):
        # find length of column i
        column_len = df[col].astype(str).str.len().max()
        # Setting the length if the column header is larger than the max column value length
        column_len = max(column_len, len(col)) + 2
        worksheet.set_column(i, i, column_len)

### Excel formatting
workbook = writer.book
cellformat = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})
worksheet = writer.sheets[sheetname1]
worksheet.set_column('H:H', cell_format=cellformat)
worksheet.set_column('J:J', cell_format=cellformat)
auto_width_column(df1)

worksheet = writer.sheets[sheetname2]
worksheet.set_column('M:M', cell_format=cellformat)
auto_width_column(df2)

worksheet = writer.sheets[sheetname3]
auto_width_column(df3)

### Save the Excel Workbook
writer.save()

### Email sending code
email_user = 'saurabh.jain@ddslab.com'
email_password = 'J@in2018oct'
email_list = ['saurabh.jain@ddslab.com']
email_send = ', '.join(email_list)

subject = 'Heartland Monthly Reporting ' + lastMonth

msg = MIMEMultipart()
msg['From'] = email_user
msg['To'] = email_send
msg['Subject'] = subject

body = '''Hey Kelly,

Please find attached the Heartland Monthly detailed report for {lastMonth}.

Kindly let me know in case of any concerns.

Thank you'''.format(lastMonth=lastMonth)

msg.attach(MIMEText(body, 'plain'))

file_name = filename
attachment = open(file_name, 'rb')

part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= " + file_name)

msg.attach(part)
text = msg.as_string()
server = smtplib.SMTP('smtp-mail.outlook.com', 587)
server.starttls()
server.login(email_user, email_password)

server.sendmail(email_user, email_list, text)
server.quit()

print('Task Completed')
