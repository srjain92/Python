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

today = datetime.date.today()
first = today.replace(day=1)
lastMonth = first - datetime.timedelta(days=1)
lastMonth = (lastMonth.strftime("%m_%Y"))

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=dds-sql03;'
                      'Database=DLPLUS;'
                      'Trusted_Connection=yes;')

script1 = ('''SELECT   
            MONTH(C.DATEINVOICED) AS 'MONTH'  
            ,YEAR(C.DATEINVOICED) AS 'YEAR'  
            ,CP.PROD_FAC
            ,COUNT(DISTINCT C.[CaseID]) AS CaseCount  
            FROM [DLPLUS].[dbo].[Cases] C (NOLOCK)  
            INNER JOIN [DLPLUS].[dbo].CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID  
            WHERE  
            MONTH(C.DATEINVOICED) = MONTH(DATEADD(M,-1,GETDATE()))
            AND YEAR(C.DATEINVOICED) = YEAR(DATEADD(M,-1,GETDATE()))   
            AND C.TYPE = 0  
            GROUP BY  
            MONTH(C.DATEINVOICED)
            ,YEAR(C.DATEINVOICED) 
            ,CP.PROD_FAC''')

script2_1 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
                DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)

            SELECT  
                MONTH(C.DATEINVOICED) AS [MONTH]
                ,YEAR(C.DATEINVOICED)  AS [YEAR]  
                ,INV.CASECOUNT AS 'CasesInvoiced'  
                ,REC.CASECOUNT AS 'CasesReceived'  
                ,UN.UNITS AS 'UnitsInvoiced'  
                ,UN.UNITS/INV.CASECOUNT as 'UnitsPerCase'  
                ,ODA.UniqueDoctors  
                ,ODA.UniqueOffices
                ,CAST(INV.CASECOUNT AS FLOAT)/CAST(ODA.UniqueOffices AS FLOAT) [Case Per Office]
                ,CAST(INV.CASECOUNT AS FLOAT)/CAST(ODA.UniqueDoctors AS FLOAT) [Case Per Dentist]
            FROM CASES C (NOLOCK)  
                LEFT JOIN (SELECT  
                             COUNT(C.CASEID) AS CASECOUNT  
                            ,MONTH(C.DATEINVOICED) AS [MONTH]  
                            ,YEAR(C.DATEINVOICED) AS [YEAR]
                        FROM CASES C (NOLOCK)  
                             INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
                       WHERE  
                             C.DATEINVOICED >= @StartDate  
                         AND C.DATEINVOICED < @EndDate 
                         AND C.[TYPE] = 0  
                       GROUP BY  
                             MONTH(C.DATEINVOICED)
                            ,YEAR(C.DATEINVOICED)) AS INV ON INV.[MONTH] = MONTH(C.DATEINVOICED) AND INV.YEAR = YEAR(C.DATEINVOICED)
                LEFT JOIN (SELECT  
                             COUNT(C.CASEID) AS CASECOUNT  
                            ,MONTH(C.DATEIN) AS [MONTH]   
                            ,YEAR(C.DATEIN) AS [YEAR]  
                        FROM CASES C (NOLOCK)  
                             INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
                       WHERE  
                             C.DATEIN >= @StartDate  
                         AND C.DATEIN < @EndDate
                         AND C.[TYPE] = 0  
                       GROUP BY  
                             MONTH(C.DATEIN)   
                            ,YEAR(C.DATEIN)) AS REC ON REC.[MONTH] = MONTH(C.DATEINVOICED) AND REC.YEAR = YEAR(C.DATEINVOICED)
                LEFT JOIN (SELECT  
                             MONTH(C.DATEINVOICED) AS [MONTH] 
                            ,YEAR(C.DATEINVOICED) AS [YEAR] 
                            ,SUM(QUANTITY) AS UNITS  
                        FROM CASEPRODUCTS CP (NOLOCK)  
                             INNER JOIN CASES C (NOLOCK) ON C.CASEID = CP.CASEID  
                             INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
                             INNER JOIN Products P (NOLOCK) ON P.ProductID = CP.ProductID  
                       WHERE  
                             C.DATEINVOICED >= @StartDate  
                         AND C.DATEINVOICED < @EndDate
                         AND C.[TYPE] = 0  
                         AND P.MAJORGROUPID IN ('ORTHO','FIXED','REMOVABLE')  
                         AND P.PRODUCTTYPEID = 'Product'  
                       GROUP BY  
                             MONTH(C.DATEINVOICED) 
                            ,YEAR(C.DATEINVOICED)) AS UN ON UN.[MONTH] = MONTH(C.DATEINVOICED) AND UN.[YEAR] = YEAR(C.DATEINVOICED) 
                LEFT JOIN (SELECT   
                             MONTH(C.DATEINVOICED) AS [MONTH] 
                            ,YEAR(C.DATEINVOICED) AS [YEAR]   
                            ,COUNT(DISTINCT D.FIRSTNAME + D.LASTNAME + D.BILLTO) AS UniqueDoctors  
                            ,COUNT(DISTINCT D.PracticeName) AS UniqueOffices  
                        FROM CASES C (NOLOCK)  
                             INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
                       WHERE  
                             C.DATEINVOICED >= @StartDate  
                         AND C.DATEINVOICED < @EndDate 
                         AND C.[TYPE] = 0  
                         AND D.BILLTO IS NOT NULL  
                       GROUP BY  
                             MONTH(C.DATEINVOICED)   
                            ,YEAR(C.DATEINVOICED) 
                      HAVING  
                             COUNT(C.CASEID) > 0) AS ODA ON ODA.[MONTH] = MONTH(C.DATEINVOICED) AND ODA.[YEAR] = YEAR(C.DATEINVOICED)  
            WHERE  
                C.[TYPE] = 0  
                AND C.DATEINVOICED >= @StartDate 
                AND C.DATEINVOICED < @EndDate 
            GROUP BY  
                MONTH(C.DATEINVOICED)
                ,YEAR(C.DATEINVOICED)  
                ,INV.CASECOUNT  
                ,REC.CASECOUNT  
                ,UN.UNITS  
                ,ODA.UniqueDoctors  
                ,ODA.UniqueOffices 
            ORDER BY
                [YEAR]
                ,[MONTH]''')

script2_2 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
                DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)

            SELECT 
                COALESCE(DFG.[DailyFlashGrouping],'Other') AS [Description] 
                ,(SUM(CP.TOTALCHARGE) - SUM(COALESCE(CP.TOTALTAX,0)))/1000 AS [Revenue]
            FROM CASEPRODUCTS CP (NOLOCK) 
                INNER JOIN  CASES C (NOLOCK) ON C.CASEID = CP.CASEID 
                INNER JOIN PRODUCTS P (NOLOCK) ON P.PRODUCTID = CP.PRODUCTID 
                INNER JOIN [UnitGroups] UG (NOLOCK) ON UG.[UnitGroupID] = P.[UnitGroupID] 
                LEFT JOIN  [DDSNet_DailyFlashProductClassGrouping] DFG ON DFG.[UnitGroupID] = P.[UnitGroupID] 
                LEFT JOIN (SELECT 
                                [TypeValue] 
                                ,MONTH([ApplyDate]) AS 'Month' 
                                ,SUM([BudgetValue]) AS 'Budget' 
                            FROM [DLPLUS].[dbo].[DDSNet_DailyFlashBudgets]  (NOLOCK)
                            WHERE 
                                MONTH([ApplyDate]) = MONTH(@StartDate)
                                AND YEAR([ApplyDate]) = YEAR(@StartDate)
                                AND [TypeID] = 'ProductClass' 
                            GROUP BY 
                                [TypeValue] 
                                ,MONTH([ApplyDate])) B ON B.[Month] = MONTH(C.DATEINVOICED) AND B.[TypeValue] = COALESCE(DFG.[DailyFlashGrouping],'Other') 
            WHERE 
                MONTH(C.DATEINVOICED) = MONTH(@StartDate)
                AND YEAR(C.DATEINVOICED) = YEAR(@StartDate)
                AND C.DATEINVOICED < @EndDate
                AND C.TYPE = 0 
            GROUP BY 
                COALESCE(DFG.[DailyFlashGrouping],'Other') 
                ,COALESCE(B.Budget,0)  
            ORDER BY 
                COALESCE(DFG.[DailyFlashGrouping],'Other')''')

script2_3 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
                DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)
                SET NOCOUNT ON;

            DECLARE @CustomerDataTable1 TABLE  
                    (  
                      [CustomerName] Varchar(60),   
                      [Value] decimal(16,3)  
                    )  
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'Other Accounts' AS [CustomerName]
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value]
                  
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate  
                AND C.[TYPE] = 0  
                AND D.CustomerTypeID <> 'UNITED_CONCORDIA'   
                AND D.BILLTO <> 'COAST_DENTAL10000'   
                AND D.BILLTO <> 'SMILECAREMAINACCOUNT'   
                AND D.BILLTO <> 'BRIGHTNOW!32028701'   
                AND D.BILLTO <> 'HEARTLAND_DE32014775'   
                AND D.BILLTO <> 'DENTALONEPARTNERSMAI'   
                AND D.BILLTO <> 'WESTERNDENTALMAIN'   
                AND D.BILLTO <> 'DENTAL_CARE_32016138'   
                AND D.BILLTO <> 'DECAMAIN'   
                AND D.BILLTO <> 'SAGEDENTALCORPMAIN'   
                AND D.BILLTO <> 'NAMERICANDNTLGRPMAIN' 
                AND D.BILLTO <> 'FAMILIADENTALGRPMAIN'  
            
            
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'Coast_SmileCare' AS [CustomerName] 
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value]
                  
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate   
                AND C.[TYPE] = 0  
                AND (D.BILLTO = 'COAST_DENTAL10000' or D.BILLTO = 'SMILECAREMAINACCOUNT')  
            
            
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'BrightNow' AS [CustomerName]   
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value]   
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate    
                AND C.[TYPE] = 0  
                AND D.BILLTO = 'BRIGHTNOW!32028701'   
            
            
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'Heartland Dental Care' AS [CustomerName]  
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value]  
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate  
                AND C.[TYPE] = 0  
                AND D.BILLTO = 'HEARTLAND_DE32014775'   
            
            
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'Dental One Partners' AS [CustomerName] 
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value] 
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate    
                AND C.[TYPE] = 0  
                AND D.BILLTO = 'DENTALONEPARTNERSMAI'  
            
            
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'Western Dental Services' AS [CustomerName] 
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value]  
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate    
                AND C.[TYPE] = 0  
                AND D.BILLTO = 'WESTERNDENTALMAIN'  
            
            
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'Dental Care Alliance' AS [CustomerName]  
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value]
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate  
                AND C.[TYPE] = 0  
                AND D.BILLTO = 'DENTAL_CARE_32016138'  
                                  
            
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'DECA' AS [CustomerName]  
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value]   
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate  
                AND C.[TYPE] = 0  
                AND D.BILLTO = 'DECAMAIN'  
            
            
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'Gentle Dental' AS [CustomerName]  
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value] 
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate  
                AND C.[TYPE] = 0  
                AND D.BILLTO = 'SAGEDENTALCORPMAIN'  
                                  
            
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'North American Dental Group' AS [CustomerName]  
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value] 
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate  
                AND C.[TYPE] = 0  
                AND D.BILLTO = 'NAMERICANDNTLGRPMAIN'
            
            
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'Familia Dental' AS [CustomerName] 
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value] 
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate    
                AND C.[TYPE] = 0  
                AND D.BILLTO = 'FAMILIADENTALGRPMAIN'   
                                  
                              
            INSERT INTO @CustomerDataTable1  
            ([CustomerName],[Value])  
            
            SELECT  
                'United Concordia' AS [CustomerName] 
                ,(SUM(C.TOTALCHARGE) - SUM(COALESCE(C.TOTALTAX,0)))/1000 AS [Value]   
            FROM CASES C (NOLOCK)   
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID  
            WHERE  
                MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END
                AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END
                AND C.DATEINVOICED < @EndDate   
                AND C.[TYPE] = 0  
                AND D.customertypeid = 'UNITED_CONCORDIA'
                
                
            SELECT  
                   [CustomerName]  
                  ,[Value] 
              FROM @CustomerDataTable1 CDT   
            --       LEFT JOIN (SELECT   
            --                        [TypeValue]  
            --                       ,SUM([BudgetValue]) AS 'Budget'  
            --                   FROM [DLPLUS].[dbo].[DDSNet_DailyFlashBudgets]   (NOLOCK)
            --                  WHERE  
            --                        MONTH([ApplyDate]) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END  
            --                    AND YEAR([ApplyDate]) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END  
            --                    AND [TypeID] = 'Customer'  
            --                  GROUP BY  
            --                        [TypeValue]) AS BG ON BG.[TypeValue] = CDT.[CustomerName]''')

script2_4 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
                DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)

                SELECT   
                    COALESCE(DFLG.[DailyFlashGrouping],'DDS US') AS 'Lab' 
                    ,COUNT(DISTINCT C.[CaseID]) AS CaseCount  
                FROM [Cases] C (NOLOCK)  
                    INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID  
                    LEFT JOIN [DDSNet_DailyFlashLabGrouping] DFLG (NOLOCK) ON DFLG.[Prod_Fac] = CP.[Prod_Fac]  
                WHERE  
                    MONTH(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN 12 ELSE MONTH(@EndDate) - 1 END   
                    AND YEAR(C.DATEINVOICED) = CASE WHEN MONTH(@EndDate) = 1 THEN YEAR(@EndDate) - 1 ELSE YEAR(@EndDate) END    
                    AND C.DATEINVOICED < @EndDate   
                    AND C.TYPE = 0  
                GROUP BY  
                    COALESCE(DFLG.[DailyFlashGrouping],'DDS US')  
                    ,COALESCE(DFLG.[Position],2)
                ORDER BY 
                    COALESCE(DFLG.[Position],2)''')

script2_5 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
                DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)

                SELECT   
                    COALESCE(DFLG.[DailyFlashGrouping],'DDS US') AS 'Lab' 
                    ,COUNT(DISTINCT C.[CaseID]) AS CaseCount  
                    ,YEAR(C.DATEINVOICED) AS 'YEAR'
                    ,MONTH(C.DATEINVOICED) AS 'MONTH'
                FROM [Cases] C (NOLOCK)  
                    INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID  
                    LEFT JOIN Products P (NOLOCK) ON P.ProductID = CP.ProductID
                    LEFT JOIN [DDSNet_DailyFlashLabGrouping] DFLG (NOLOCK) ON DFLG.[Prod_Fac] = CP.[Prod_Fac]  
                WHERE  
                    C.DATEINVOICED >= @StartDate
                    AND C.DATEINVOICED < @EndDate   
                    AND C.TYPE = 0  
                    AND P.MAJORGROUPID = 'Fixed'
                GROUP BY  
                    YEAR(C.DATEINVOICED)
                    ,MONTH(C.DATEINVOICED)
                    ,COALESCE(DFLG.[DailyFlashGrouping],'DDS US')  
                    ,COALESCE(DFLG.[Position],2)
                ORDER BY 
                    YEAR(C.DATEINVOICED)
                    ,MONTH(C.DATEINVOICED)
                    ,COALESCE(DFLG.[Position],2)''')

script2_6 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
                DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)

                SELECT   
                    COALESCE(DFLG.[DailyFlashGrouping],'DDS US') AS 'Lab' 
                    ,COUNT(DISTINCT C.[CaseID]) AS CaseCount  
                    ,YEAR(C.DATEINVOICED) AS 'YEAR'
                    ,MONTH(C.DATEINVOICED) AS 'MONTH'
                FROM [Cases] C (NOLOCK)  
                    INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID  
                    LEFT JOIN Products P (NOLOCK) ON P.ProductID = CP.ProductID
                    LEFT JOIN [DDSNet_DailyFlashLabGrouping] DFLG (NOLOCK) ON DFLG.[Prod_Fac] = CP.[Prod_Fac]  
                WHERE  
                    C.DATEINVOICED >= @StartDate
                    AND C.DATEINVOICED < @EndDate   
                    AND C.TYPE = 0  
                    AND P.MAJORGROUPID = 'Removable'
                GROUP BY  
                    YEAR(C.DATEINVOICED)
                    ,MONTH(C.DATEINVOICED)
                    ,COALESCE(DFLG.[DailyFlashGrouping],'DDS US')
                    ,COALESCE(DFLG.[Position],2)
                ORDER BY 
                    YEAR(C.DATEINVOICED)
                    ,MONTH(C.DATEINVOICED)
                    ,COALESCE(DFLG.[Position],2)''')

script2_7 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
                DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)
                
                SELECT 
                    COALESCE(DFLG.[DailyFlashGrouping],'DDS US') AS 'Lab' 
                    ,COUNT(DISTINCT C.[CaseID]) AS CaseCount  
                    ,YEAR(C.DATEINVOICED) AS 'YEAR'
                    ,MONTH(C.DATEINVOICED) AS 'MONTH'
                FROM [Cases] C (NOLOCK)  
                    INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID  
                    LEFT JOIN Products P (NOLOCK) ON P.ProductID = CP.ProductID
                    LEFT JOIN [DDSNet_DailyFlashLabGrouping] DFLG (NOLOCK) ON DFLG.[Prod_Fac] = CP.[Prod_Fac]  
                WHERE  
                    C.DATEINVOICED >= @StartDate
                    AND C.DATEINVOICED < @EndDate   
                    AND C.TYPE = 0  
                    AND P.MAJORGROUPID = 'Ortho'
                GROUP BY  
                    YEAR(C.DATEINVOICED)
                    ,MONTH(C.DATEINVOICED)
                    ,COALESCE(DFLG.[DailyFlashGrouping],'DDS US')  
                    ,COALESCE(DFLG.[Position],2)
                ORDER BY 
                    YEAR(C.DATEINVOICED)
                    ,MONTH(C.DATEINVOICED)
                    ,COALESCE(DFLG.[Position],2)''')

script3 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
              DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)

            SELECT 
                YEAR(C.DateInvoiced) AS YEAR
                ,MONTH(C.DateInvoiced) AS MONTH
                ,COUNT(DISTINCT C.CASEID) AS CaseCount
                ,COUNT(DISTINCT CASE WHEN COALESCE(NCH.TotalCharge, 0) < 10 THEN NCH.CaseID ELSE NULL END) 'NoChargeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL THEN C.CaseID ELSE NULL END) 'RemakeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 THEN C.CaseID ELSE NULL END) 'NoChargeRemakeCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL THEN C.CaseID ELSE NULL END) 'WarrantyCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL AND COALESCE(NCH.TotalCharge, 0) < 10 THEN NCH.CaseID ELSE NULL END) 'NoChargeWarrantyCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL AND C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 THEN C.CaseID ELSE NULL END) 'NoChargeWarrantyRemakeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 AND W.CASEID IS NULL THEN C.CaseID ELSE NULL END) 'NoChargeRemakesExWarranty'
            FROM CASES C 
                INNER JOIN DOCTORS D ON D.DOCTORID = C.DOCTORID
                INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID    
                INNER JOIN Products P (NOLOCK) ON CP.ProductID = P.ProductID    
                LEFT JOIN (SELECT 
                                CP.[CaseID]
                            FROM [DLPlus].[dbo].[CaseProducts] CP
                                INNER JOIN CASES C ON C.CASEID = CP.CASEID
                                INNER JOIN Doctors D on D.DoctorID = C.DoctorID
                            WHERE 
                                C.DateInvoiced >= @StartDate
                                AND C.DateInvoiced < @EndDate
                                AND PRODUCTID IN (SELECT 
                                                    [ProductID]
                                                  FROM [DLPlus].[dbo].[Products]
                                                  WHERE
                                                    [ProductID] like '%warranty%'
                                                    OR [Description] like '%warranty%')
                            GROUP BY
                                CP.[CaseID]) AS W ON W.CASEID = C.CASEID
                LEFT JOIN (SELECT 
                                CP.[CaseID]
                            FROM [DLPlus].[dbo].[CaseProducts] CP
                                INNER JOIN CASES C ON C.CASEID = CP.CASEID
                                INNER JOIN Doctors D on D.DoctorID = C.DoctorID
                            WHERE 
                                C.DateInvoiced >= @StartDate
                                AND C.DateInvoiced < @EndDate
                                AND PRODUCTID IN (SELECT 
                                                    [ProductID]
                                                  FROM [DLPlus].[dbo].[Products]
                                                  
                                                     WHERE DESCRIPTION LIKE '%$%' OR PRODUCTID LIKE 'cp_%' OR PRODUCTID LIKE 'COUPON_%')
                            GROUP BY
                                CP.[CaseID]) AS CO ON CO.CASEID = C.CASEID
                LEFT JOIN (SELECT A.CaseID
                                  ,A.MONTH
                                  ,A.YEAR
                                  ,A.CaseTotalCharge
                                  ,SUM(A.ProductTotalCharge) TotalCharge
                              FROM (SELECT C.CaseID
                                          ,P.ProductID
                                          ,C.TotalCharge CaseTotalCharge
                                          ,CP.TotalCharge ProductTotalCharge
                                          ,DL.PriceListID
                                          ,MONTH(C.DateInvoiced) MONTH
                                          ,YEAR(DateInvoiced) YEAR
                                      FROM CASES C 
                                           INNER JOIN DOCTORS D ON D.DOCTORID = C.DOCTORID
                                           INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID    
                                           INNER JOIN Products P (NOLOCK) ON CP.ProductID = P.ProductID
                                           LEFT JOIN DocLab DL ON D.doctorid = DL.DoctorID 
                                     WHERE C.DateInvoiced >= @StartDate
                                       AND C.DateInvoiced < @EndDate
                                       AND [TYPE] = 0
                                       AND P.ProductID NOT IN (SELECT ProductID
                                                                 FROM ProductPrices
                                                                WHERE PriceListID = DL.PriceListID 
                                                                  AND ZeroPriced = 1)
                                     GROUP BY C.CaseID
                                          ,P.ProductID
                                          ,CP.TotalCharge
                                          ,C.TotalCharge
                                          ,DL.PriceListID
                                          ,MONTH(C.DateInvoiced) 
                                          ,YEAR(DateInvoiced)) A
                             GROUP BY A.CaseID
                                  ,A.MONTH
                                  ,A.YEAR
                                  ,A.CaseTotalCharge) NCH ON NCH.CaseID = C.CaseID
            WHERE
                C.DateInvoiced >= @StartDate
                AND C.DateInvoiced < @EndDate
                AND [TYPE] = 0
                AND P.MajorGroupID <> 'Shipping' 
            GROUP BY
                YEAR(C.DateInvoiced)
                ,MONTH(C.DateInvoiced)
            ORDER BY
                YEAR(C.DateInvoiced) 
                ,MONTH(C.DateInvoiced)''')

script4 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
                DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)

            SELECT 
                YEAR(C.DateInvoiced) AS YEAR
                ,MONTH(C.DateInvoiced) AS MONTH
                ,P.MajorGroupID
                ,COUNT(DISTINCT C.CASEID) AS CaseCount
                ,COUNT(DISTINCT CASE WHEN COALESCE(NCH.TotalCharge, 0) < 10 THEN NCH.CaseID ELSE NULL END) 'NoChargeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL THEN C.CaseID ELSE NULL END) 'RemakeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 THEN C.CaseID ELSE NULL END) 'NoChargeRemakeCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL THEN C.CaseID ELSE NULL END) 'WarrantyCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL AND COALESCE(NCH.TotalCharge, 0) < 10 THEN NCH.CaseID ELSE NULL END) 'NoChargeWarrantyCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL AND C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 THEN C.CaseID ELSE NULL END) 'NoChargeWarrantyRemakeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 AND W.CASEID IS NULL THEN C.CaseID ELSE NULL END) 'NoChargeRemakesExWarranty'
            FROM CASES C  (NOLOCK)
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID
                INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID    
                INNER JOIN Products P (NOLOCK) ON CP.ProductID = P.ProductID    
                LEFT JOIN (SELECT 
                                CP.[CaseID]
                            FROM [DLPlus].[dbo].[CaseProducts] CP (NOLOCK)
                                INNER JOIN CASES C (NOLOCK) ON C.CASEID = CP.CASEID
                                INNER JOIN Doctors D (NOLOCK) on D.DoctorID = C.DoctorID
                            WHERE 
                                C.DateInvoiced >= @StartDate
                                AND C.DateInvoiced < @EndDate
                                AND PRODUCTID IN (SELECT 
                                                    [ProductID]
                                                  FROM [DLPlus].[dbo].[Products] (NOLOCK)
                                                  WHERE
                                                    [ProductID] like '%warranty%'
                                                    OR [Description] like '%warranty%')
                            GROUP BY
                                CP.[CaseID]) AS W ON W.CASEID = C.CASEID
                LEFT JOIN (SELECT 
                                CP.[CaseID]
                            FROM [DLPlus].[dbo].[CaseProducts] CP (NOLOCK)
                                INNER JOIN CASES C (NOLOCK) ON C.CASEID = CP.CASEID
                                INNER JOIN Doctors D (NOLOCK) on D.DoctorID = C.DoctorID
                            WHERE 
                                C.DateInvoiced >= @StartDate
                                AND C.DateInvoiced < @EndDate
                                AND PRODUCTID IN (SELECT 
                                                    [ProductID]
                                                  FROM [DLPlus].[dbo].[Products] (NOLOCK)
                                                  
                                                     WHERE DESCRIPTION LIKE '%$%' OR PRODUCTID LIKE 'cp_%' OR PRODUCTID LIKE 'COUPON_%')
                            GROUP BY
                                CP.[CaseID]) AS CO ON CO.CASEID = C.CASEID
                LEFT JOIN (SELECT A.CaseID
                                  ,A.MONTH
                                  ,A.YEAR
                                  ,A.CaseTotalCharge
                                  ,SUM(A.ProductTotalCharge) TotalCharge
                              FROM (SELECT C.CaseID
                                          ,P.ProductID
                                          ,C.TotalCharge CaseTotalCharge
                                          ,CP.TotalCharge ProductTotalCharge
                                          ,DL.PriceListID
                                          ,MONTH(C.DateInvoiced) MONTH
                                          ,YEAR(DateInvoiced) YEAR
                                      FROM CASES C  (NOLOCK)
                                           INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID
                                           INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID    
                                           INNER JOIN Products P (NOLOCK) ON CP.ProductID = P.ProductID
                                           LEFT JOIN DocLab DL (NOLOCK) ON D.doctorid = DL.DoctorID 
                                     WHERE C.DateInvoiced >= @StartDate
                                       AND C.DateInvoiced < @EndDate
                                       AND [TYPE] = 0
                                       AND P.ProductID NOT IN (SELECT ProductID
                                                                 FROM ProductPrices (NOLOCK)
                                                                WHERE PriceListID = DL.PriceListID 
                                                                  AND ZeroPriced = 1)
                                     GROUP BY C.CaseID
                                          ,P.ProductID
                                          ,CP.TotalCharge
                                          ,C.TotalCharge
                                          ,DL.PriceListID
                                          ,MONTH(C.DateInvoiced) 
                                          ,YEAR(DateInvoiced)) A
                             GROUP BY A.CaseID
                                  ,A.MONTH
                                  ,A.YEAR
                                  ,A.CaseTotalCharge) NCH ON NCH.CaseID = C.CaseID
                                
            WHERE
                C.DateInvoiced >= @StartDate
                AND C.DateInvoiced < @EndDate
                AND [TYPE] = 0
                AND P.MajorGroupID <> 'Shipping' 
            GROUP BY
                YEAR(C.DateInvoiced)
                ,MONTH(C.DateInvoiced)
                ,P.MajorGroupID
            ORDER BY
                YEAR(C.DateInvoiced) 
                ,MONTH(C.DateInvoiced)
                ,P.MajorGroupID''')

script5 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
              DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)

            SELECT 
                YEAR(C.DateInvoiced) AS YEAR
                ,MONTH(C.DateInvoiced) AS MONTH
                --,COALESCE(BT.DOCTORID,D.DOCTORID) AS AccountID
                ,COALESCE(BT.PRACTICENAME,'') AS Account
                ,COALESCE(BT.CUSTOMERTYPEID,D.CUSTOMERTYPEID,'') AS CustomerType
                ,ACCTREV.Revenue
                ,COUNT(DISTINCT C.CASEID) AS CaseCount
                ,COUNT(DISTINCT CASE WHEN COALESCE(NCH.TotalCharge, 0) < 10 THEN NCH.CaseID ELSE NULL END) 'NoChargeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL THEN C.CaseID ELSE NULL END) 'RemakeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 THEN C.CaseID ELSE NULL END) 'NoChargeRemakeCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL THEN C.CaseID ELSE NULL END) 'WarrantyCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL AND COALESCE(NCH.TotalCharge, 0) < 10 THEN NCH.CaseID ELSE NULL END) 'NoChargeWarrantyCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL AND C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 THEN C.CaseID ELSE NULL END) 'NoChargeWarrantyRemakeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 AND W.CASEID IS NULL THEN C.CaseID ELSE NULL END) 'NoChargeRemakesExWarranty'
            FROM CASES C  (NOLOCK)
                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID
                LEFT JOIN DOCTORS BT (NOLOCK) ON BT.DOCTORID = D.BILLTO
                INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID    
                INNER JOIN Products P (NOLOCK) ON CP.ProductID = P.ProductID    
                LEFT JOIN (SELECT 
                                CP.[CaseID]
                            FROM [DLPlus].[dbo].[CaseProducts] CP (NOLOCK)
                                INNER JOIN CASES C (NOLOCK) ON C.CASEID = CP.CASEID
                                INNER JOIN Doctors D (NOLOCK) on D.DoctorID = C.DoctorID
                            WHERE 
                                C.DateInvoiced >= @StartDate
                                AND C.DateInvoiced < @EndDate
                                AND C.[TYPE] = 0
                                AND PRODUCTID IN (SELECT 
                                                    [ProductID]
                                                  FROM [DLPlus].[dbo].[Products] (NOLOCK)
                                                  WHERE
                                                    [ProductID] like '%warranty%'
                                                    OR [Description] like '%warranty%')
                            GROUP BY
                                CP.[CaseID]) AS W ON W.CASEID = C.CASEID
                LEFT JOIN (SELECT 
                                CP.[CaseID]
                            FROM [DLPlus].[dbo].[CaseProducts] CP (NOLOCK)
                                INNER JOIN CASES C (NOLOCK) ON C.CASEID = CP.CASEID
                                INNER JOIN Doctors D (NOLOCK) on D.DoctorID = C.DoctorID
                            WHERE 
                                C.DateInvoiced >= @StartDate
                                AND C.DateInvoiced < @EndDate
                                AND C.[TYPE] = 0
                                AND PRODUCTID IN (SELECT 
                                                    [ProductID]
                                                  FROM [DLPlus].[dbo].[Products] (NOLOCK)
                                                  
                                                     WHERE DESCRIPTION LIKE '%$%' OR PRODUCTID LIKE 'cp_%' OR PRODUCTID LIKE 'COUPON_%')
                            GROUP BY
                                CP.[CaseID]) AS CO ON CO.CASEID = C.CASEID
                LEFT JOIN (SELECT
                                YEAR(C.DateInvoiced) AS YEAR
                                ,MONTH(C.DateInvoiced) AS MONTH
                                ,COALESCE(BT.PRACTICENAME,'') as Account
                                ,SUM(COALESCE(C.TOTALCHARGE,0)) - SUM(COALESCE(C.TOTALTAX,0)) AS REVENUE
                            FROM CASES C (NOLOCK) 
                                INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID
                                LEFT JOIN DOCTORS BT (NOLOCK) ON BT.DOCTORID = D.BILLTO
                            WHERE
                                C.DateInvoiced >= @StartDate
                                AND C.DateInvoiced < @EndDate
                                AND C.TYPE = 0
                            GROUP BY
                                YEAR(C.DateInvoiced) 
                                ,MONTH(C.DateInvoiced)
                                ,COALESCE(BT.PRACTICENAME,'')) AS ACCTREV ON ACCTREV.Account = COALESCE(BT.PRACTICENAME,'') AND ACCTREV.YEAR = YEAR(C.DateInvoiced) AND ACCTREV.MONTH = MONTH(C.DateInvoiced)
                LEFT JOIN (SELECT A.CaseID
                                  ,A.MONTH
                                  ,A.YEAR
                                  ,A.CaseTotalCharge
                                  ,SUM(A.ProductTotalCharge) TotalCharge
                              FROM (SELECT C.CaseID
                                          ,P.ProductID
                                          ,C.TotalCharge CaseTotalCharge
                                          ,CP.TotalCharge ProductTotalCharge
                                          ,DL.PriceListID
                                          ,MONTH(C.DateInvoiced) MONTH
                                          ,YEAR(DateInvoiced) YEAR
                                      FROM CASES C  (NOLOCK)
                                           INNER JOIN DOCTORS D (NOLOCK) ON D.DOCTORID = C.DOCTORID
                                           INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID    
                                           INNER JOIN Products P (NOLOCK) ON CP.ProductID = P.ProductID
                                           LEFT JOIN DocLab DL (NOLOCK) ON D.doctorid = DL.DoctorID 
                                     WHERE C.DateInvoiced >= @StartDate
                                       AND C.DateInvoiced < @EndDate
                                       AND [TYPE] = 0
                                       AND P.ProductID NOT IN (SELECT ProductID
                                                                 FROM ProductPrices (NOLOCK)
                                                                WHERE PriceListID = DL.PriceListID 
                                                                  AND ZeroPriced = 1)
                                     GROUP BY C.CaseID
                                          ,P.ProductID
                                          ,CP.TotalCharge
                                          ,C.TotalCharge
                                          ,DL.PriceListID
                                          ,MONTH(C.DateInvoiced) 
                                          ,YEAR(DateInvoiced)) A
                             GROUP BY A.CaseID
                                  ,A.MONTH
                                  ,A.YEAR
                                  ,A.CaseTotalCharge) NCH ON NCH.CaseID = C.CaseID
                                
            WHERE
                CONVERT(DATE, C.DateInvoiced) >= @StartDate
                AND CONVERT(DATE, C.DateInvoiced) < @EndDate
                AND [TYPE] = 0
                AND P.MajorGroupID <> 'Shipping' 
                AND BT.PRACTICENAME IN ('Heartland Dental Care',
            'Brightnow! - Main Acct',
            'Coast Dental',
            'Dental One Partners',
            'Western Dental Services, Inc.',
            'KOS Services, LLC',
            'DECA',
            'SmileCare',
            'Jefferson Dental',
            'Sage Dental Corporate',
            'MB2 Dental Solutions',
            'Dental Care Alliance - Main',
            'Familia Dental Group',
            'North American Dental Group',
            'Pacific Dental Services',
            '42 North Dental, LLC',
            'Spring Dental Group',
            'Perfect Teeth',
            'Benevis, LLC',
            'Lovett Dental Corp',
            'Perfect Dental Management',
            'Smile Workshop',
            'Dimensional Dental - Corp',
            'Mortenson Family Dental - Corp.',
            'Rodeo Dental')
                
            GROUP BY
                YEAR(C.DateInvoiced)
                ,MONTH(C.DateInvoiced)
                --,COALESCE(BT.DOCTORID,D.DOCTORID)
                ,COALESCE(BT.PRACTICENAME,'')
                ,COALESCE(BT.CUSTOMERTYPEID,D.CUSTOMERTYPEID,'')
                ,ACCTREV.Revenue
            ORDER BY
                YEAR(C.DateInvoiced)
                ,MONTH(C.DateInvoiced)
                --,COALESCE(BT.DOCTORID,D.DOCTORID)
                ,COALESCE(BT.PRACTICENAME,'')
                ,COALESCE(BT.CUSTOMERTYPEID,D.CUSTOMERTYPEID,'')''')

script6 = ('''DECLARE @StartDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,0,GETDATE()) -1, 0)
              DECLARE @EndDate AS DATETIME = DATEADD(MONTH, DATEDIFF(MONTH,1,GETDATE()) +0, 0)

            SELECT 
                YEAR(C.DateInvoiced) AS YEAR
                ,MONTH(C.DateInvoiced) AS MONTH
                ,COALESCE(BT.PRACTICENAME,D.PRACTICENAME) AS Account
                ,COALESCE(BT.CUSTOMERTYPEID,D.CUSTOMERTYPEID,'') AS CustomerType
                ,ACCTREV.Revenue
                ,COUNT(DISTINCT C.CASEID) AS CaseCount
                ,COUNT(DISTINCT CASE WHEN COALESCE(NCH.TotalCharge, 0) < 10 THEN NCH.CaseID ELSE NULL END) 'NoChargeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL THEN C.CaseID ELSE NULL END) 'RemakeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 THEN C.CaseID ELSE NULL END) 'NoChargeRemakeCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL THEN C.CaseID ELSE NULL END) 'WarrantyCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL AND COALESCE(NCH.TotalCharge, 0) < 10 THEN NCH.CaseID ELSE NULL END) 'NoChargeWarrantyCases'
                ,COUNT(DISTINCT CASE WHEN W.CASEID IS NOT NULL AND C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 THEN C.CaseID ELSE NULL END) 'NoChargeWarrantyRemakeCases'
                ,COUNT(DISTINCT CASE WHEN C.RemakeOf IS NOT NULL AND COALESCE(C.TotalCharge, 0) < 10 AND W.CASEID IS NULL THEN C.CaseID ELSE NULL END) 'NoChargeRemakesExWarranty'
            FROM CASES C(NOLOCK) 
                INNER JOIN DOCTORS D(NOLOCK) ON D.DOCTORID = C.DOCTORID
                LEFT JOIN DOCTORS BT(NOLOCK) ON BT.DOCTORID = D.BILLTO
                INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID    
                INNER JOIN Products P (NOLOCK) ON CP.ProductID = P.ProductID    
                LEFT JOIN (SELECT 
                                CP.[CaseID]
                            FROM [DLPlus].[dbo].[CaseProducts] CP(NOLOCK)
                                INNER JOIN CASES C(NOLOCK) ON C.CASEID = CP.CASEID
                                INNER JOIN Doctors D(NOLOCK) on D.DoctorID = C.DoctorID
                            WHERE 
                                C.DateInvoiced >= @StartDate
                                AND C.DateInvoiced < @EndDate
                                AND C.[TYPE] = 0
                                AND PRODUCTID IN (SELECT 
                                                    [ProductID]
                                                  FROM [DLPlus].[dbo].[Products](NOLOCK)
                                                  WHERE
                                                    [ProductID] like '%warranty%'
                                                    OR [Description] like '%warranty%')
                            GROUP BY
                                CP.[CaseID]) AS W ON W.CASEID = C.CASEID
                LEFT JOIN (SELECT 
                                CP.[CaseID]
                            FROM [DLPlus].[dbo].[CaseProducts] CP(NOLOCK)
                                INNER JOIN CASES C(NOLOCK) ON C.CASEID = CP.CASEID
                                INNER JOIN Doctors D(NOLOCK) on D.DoctorID = C.DoctorID
                            WHERE 
                                C.DateInvoiced >= @StartDate
                                AND C.DateInvoiced < @EndDate
                                AND C.[TYPE] = 0
                                AND PRODUCTID IN (SELECT 
                                                    [ProductID]
                                                  FROM [DLPlus].[dbo].[Products](NOLOCK)
                                                  
                                                     WHERE DESCRIPTION LIKE '%$%' OR PRODUCTID LIKE 'cp_%' OR PRODUCTID LIKE 'COUPON_%')
                            GROUP BY
                                CP.[CaseID]) AS CO ON CO.CASEID = C.CASEID
                LEFT JOIN (SELECT
                                YEAR(C.DateInvoiced) AS YEAR
                                ,MONTH(C.DateInvoiced) AS MONTH
                                ,COALESCE(BT.PRACTICENAME,'') as Account
                                ,SUM(COALESCE(C.TOTALCHARGE,0)) - SUM(COALESCE(C.TOTALTAX,0)) AS REVENUE
                            FROM CASES C (NOLOCK)
                                INNER JOIN DOCTORS D(NOLOCK) ON D.DOCTORID = C.DOCTORID
                                LEFT JOIN DOCTORS BT(NOLOCK) ON BT.DOCTORID = D.BILLTO
                            WHERE
                                C.DateInvoiced >= @StartDate
                                AND C.DateInvoiced < @EndDate
                                AND C.TYPE = 0
                            GROUP BY
                                YEAR(C.DateInvoiced) 
                                ,MONTH(C.DateInvoiced)
                                ,COALESCE(BT.PRACTICENAME,'')) AS ACCTREV ON ACCTREV.Account = COALESCE(BT.PRACTICENAME,'') AND ACCTREV.YEAR = YEAR(C.DateInvoiced) AND ACCTREV.MONTH = MONTH(C.DateInvoiced)
                LEFT JOIN (SELECT A.CaseID
                                  ,A.MONTH
                                  ,A.YEAR
                                  ,A.CaseTotalCharge
                                  ,SUM(A.ProductTotalCharge) TotalCharge
                              FROM (SELECT C.CaseID
                                          ,P.ProductID
                                          ,C.TotalCharge CaseTotalCharge
                                          ,CP.TotalCharge ProductTotalCharge
                                          ,DL.PriceListID
                                          ,MONTH(C.DateInvoiced) MONTH
                                          ,YEAR(DateInvoiced) YEAR
                                      FROM CASES C(NOLOCK) 
                                           INNER JOIN DOCTORS D(NOLOCK) ON D.DOCTORID = C.DOCTORID
                                           INNER JOIN CaseProducts CP (NOLOCK) ON C.CaseID = CP.CaseID    
                                           INNER JOIN Products P (NOLOCK) ON CP.ProductID = P.ProductID
                                           LEFT JOIN DocLab DL(NOLOCK) ON D.doctorid = DL.DoctorID 
                                     WHERE C.DateInvoiced >= @StartDate
                                       AND C.DateInvoiced < @EndDate
                                       AND [TYPE] = 0
                                       AND P.ProductID NOT IN (SELECT ProductID
                                                                 FROM ProductPrices(NOLOCK)
                                                                WHERE PriceListID = DL.PriceListID 
                                                                  AND ZeroPriced = 1)
                                     GROUP BY C.CaseID
                                          ,P.ProductID
                                          ,CP.TotalCharge
                                          ,C.TotalCharge
                                          ,DL.PriceListID
                                          ,MONTH(C.DateInvoiced) 
                                          ,YEAR(DateInvoiced)) A
                             GROUP BY A.CaseID
                                  ,A.MONTH
                                  ,A.YEAR
                                  ,A.CaseTotalCharge) NCH ON NCH.CaseID = C.CaseID
                                
            WHERE
                C.DateInvoiced >= @StartDate
                AND C.DateInvoiced < @EndDate
                AND [TYPE] = 0
                AND P.MajorGroupID <> 'Shipping' 
                AND D.CUSTOMERTYPEID IN ('STANDARD_ACCOUNTS', 'UNITED_CONCORDIA', 'GPO')
            GROUP BY
                YEAR(C.DateInvoiced)
                ,MONTH(C.DateInvoiced)
                ,COALESCE(BT.PRACTICENAME,D.PRACTICENAME)
                ,COALESCE(BT.CUSTOMERTYPEID,D.CUSTOMERTYPEID,'')
                ,ACCTREV.Revenue
            ORDER BY
                YEAR(C.DateInvoiced) 
                ,MONTH(C.DateInvoiced)
                ,COALESCE(BT.PRACTICENAME,D.PRACTICENAME)
                ,COALESCE(BT.CUSTOMERTYPEID,D.CUSTOMERTYPEID,'')
                ,ACCTREV.Revenue''')

### Create a Pandas dataframe from the data
df1 = pd.read_sql_query(script1, conn)
df2_1 = pd.read_sql_query(script2_1, conn)
df2_2 = pd.read_sql_query(script2_2, conn)
df2_3 = pd.read_sql_query(script2_3, conn)
df2_4 = pd.read_sql_query(script2_4, conn)
df2_5 = pd.read_sql_query(script2_5, conn)
df2_6 = pd.read_sql_query(script2_6, conn)
df2_7 = pd.read_sql_query(script2_7, conn)
df3 = pd.read_sql_query(script3, conn)
df4 = pd.read_sql_query(script4, conn)
df5 = pd.read_sql_query(script5, conn)
df6 = pd.read_sql_query(script6, conn)

filename = 'KPI_Monthly' + '_' + lastMonth + '.xlsx'
sheetname1 = 'Cases by Lab Data' + '_' + lastMonth
sheetname2 = 'Cases and Dentists Data' + '_' + lastMonth
sheetname3 = 'No Charge Remake Data' + '_' + lastMonth
sheetname4 = 'Remakes by Department' + '_' + lastMonth
sheetname5 = 'Top 25 Customer Data' + '_' + lastMonth
sheetname6 = 'UC Data' + '_' + lastMonth

### Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(filename, engine='xlsxwriter')

### Convert the dataframe to an XlsxWriter Excel object
df1.to_excel(writer, sheet_name=sheetname1, index=False)
df2_1.to_excel(writer, sheet_name=sheetname2, index=False)
df2_2.to_excel(writer, sheet_name=sheetname2, index=False, startrow=4)
df2_3.to_excel(writer, sheet_name=sheetname2, index=False, startrow=17)
df2_4.to_excel(writer, sheet_name=sheetname2, index=False, startrow=32)
df2_5.to_excel(writer, sheet_name=sheetname2, index=False, startrow=39)
df2_6.to_excel(writer, sheet_name=sheetname2, index=False, startrow=46)
df2_7.to_excel(writer, sheet_name=sheetname2, index=False, startrow=53)
df3.to_excel(writer, sheet_name=sheetname3, index=False)
df4.to_excel(writer, sheet_name=sheetname4, index=False)
df5.to_excel(writer, sheet_name=sheetname5, index=False)
df6.to_excel(writer, sheet_name=sheetname6, index=False)

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
worksheet = writer.sheets[sheetname1]
auto_width_column(df1)

worksheet = writer.sheets[sheetname2]
auto_width_column(df2_1)
auto_width_column(df2_3)
cellformat = workbook.add_format({'bold': True, 'font_color': 'red', 'bg_color': 'yellow'})
worksheet.write_string(38, 0, 'FIXED', cell_format=cellformat)
worksheet.write_string(45, 0, 'REMOVABLE', cell_format=cellformat)
worksheet.write_string(52, 0, 'ORTHO', cell_format=cellformat)

worksheet = writer.sheets[sheetname3]
auto_width_column(df3)
worksheet = writer.sheets[sheetname4]
auto_width_column(df4)
worksheet = writer.sheets[sheetname5]
auto_width_column(df5)
worksheet = writer.sheets[sheetname6]
auto_width_column(df6)

### Save the Excel Workbook
writer.save()

### Email sending code
email_user = 'saurabh.jain@ddslab.com'
email_password = 'J@in2018oct'
email_list = ['Jesse.Miller@ddslab.com']
email_send = ', '.join(email_list)

subject = 'KPI Monthly Data ' + lastMonth

msg = MIMEMultipart()
msg['From'] = email_user
msg['To'] = email_send
msg['Subject'] = subject

body = '''Hey Jesse,
         
Please find attached the KPI Monthly data'''

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
