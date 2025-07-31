<#
  Purpose: connect to SQL Server and run query report for sales employee
  date:20200819
#>

 $SQL_sales_employee="SELECT t.SlpCode,t.SlpName,t.Email FROM OSLP t WHERE t.Active ='Y' AND isnull(t.Email,'')<>''"
 $SQL_Report="Declare @FromDate as DateTime
Declare @ToDate as DateTime
Set @FromDate=CONVERT(varchar(12) ,dateadd(d,-day(getdate())+1,getdate()) ,111) 

Set @ToDate=getdate()

    Select   					
			Convert( varchar(10),@FromDate,23) as FromDate,		
			Convert( varchar(10),@ToDate,23) as ToDate, 		
			T0.cardcode as CustomerCode , 		
            T0.cardname as CustomerName,					
            T0.CardFName as ForeignName,					
            P1.SlpName as SalesmenCode,					
            Isnull(T1.OpenBalance,0) AS PeriodBegingBal					
            ,( isnull(T3.CurInvAmt,0)-ISNULL(T3.VatSum,0)  - isnull(T4.CurCrdAmt,0)+ISNULL(T4.VatSum,0)) -  ISNULL(T14.Freight ,0)  as PeriodNetSales 					
            ,ISNULL(T3.VatSum,0)-ISNULL(T4.VatSum,0)   AS PeriodVat					
            ,isnull(T3.CurInvAmt,0)  - isnull(T4.CurCrdAmt,0) -  ISNULL(T14.Freight ,0) As PeriodSalesWithVat					
            ,isnull( T12.CASHSUM,0)+isnull( T12.TrsfrSum,0)-Isnull(T13.TrsfrSum,0) As [Cash Collection]					
            ,isnull( T12.BoeSum,0)-Isnull(T13.BoeSum,0)  as [Draft Collection] 					
            ,ISNULL(T15.ManualJE,0)-Isnull(T13.Adjust,0) AS AdjustMent,					
            Isnull(T1.OpenBalance,0) +isnull(T3.CurInvAmt,0) - isnull(T4.CurCrdAmt,0)-ISNULL(T14.Freight ,0)
            -(isnull( T12.CASHSUM,0)+isnull( T12.TrsfrSum,0)					
            -Isnull(T13.TrsfrSum,0))-(( isnull( T12.BoeSum,0))-Isnull(T13.BoeSum,0))+
           ISNULL(T15.ManualJE,0)-Isnull(T13.Adjust,0)-isnull( T51.ReconSum,0)  as PeriodEndingBal					
				From OCRD T0 LEFT JOIN OSLP  P1 ON T0.SlpCode=P1.SlpCode	
-------------------------------------------Openning Transation-----------------------------------------------					
-------------------------------------------------------------------------------------------------------------					
Left join 					
(SELECT ShortName,Sum(Debit-Credit) as OpenBalance FROM JDT1 Where RefDate<@FromDate AND Account IN ('150110-01','150210-01','150220-01','161140-01','223240-01') 					
GROUP BY ShortName) 					
T1					
ON T0.CardCode=T1.ShortName					
		
------------------------------------------------------------------------------------------------------------					
---------------------------------------This Month Transaction-----------------------------------------------					
				
Left join 					
(SELECT CardCode,Sum(BoeSum) as BoeSum ,Sum(TrsfrSum) as TrsfrSum,Sum(CashSum) as CashSum,Sum([CheckSum]) as  [CheckSum] 					
From 					
(SELECT T1.DocNum,ShortName AS CardCode,Case When T1.BoeSum>0 Then  Sum(Credit-Debit) End As BoeSum,					
                 Case When T1.TrsfrSum>T1.CashSum+T1.[CheckSum] Then  Sum(Credit-Debit) End As TrsfrSum,					
                 Case When T1.CashSum>T1.TrsfrSum+T1.[CheckSum] Then  Sum(Credit-Debit) End As CashSum,					
                 Case When T1.[CheckSum]>T1.TrsfrSum+T1.CashSum Then  Sum(Credit-Debit) End As [CheckSum]					
					
FROM JDT1  T0 with (nolock) LEFT JOIN ORCT T1  with (nolock) ON T0.BaseRef=T1.DocEntry					
WHERE TransType='24' AND Account IN ('150110-01','150210-01','150220-01','161140-01','223240-01') AND T0.RefDate Between @FromDate and @ToDate		
and TrsfrAcct not in ('173000-02')			
GROUP BY T1.DocNum, T0.ShortName					
,T1.BoeSum,T1.TrsfrSum,T1.CashSum,T1.[CheckSum]

union all
SELECT T1.DocNum,ShortName AS CardCode,Case When T1.TrsfrSum>T1.CashSum+T1.[CheckSum] Then  Sum(Credit-Debit) End As BoeSum,
                 Case When T1.BoeSum>0 Then  Sum(Credit-Debit) End As TrsfrSum,									
                 Case When T1.CashSum>T1.TrsfrSum+T1.[CheckSum] Then  Sum(Credit-Debit) End As CashSum,					
                 Case When T1.[CheckSum]>T1.TrsfrSum+T1.CashSum Then  Sum(Credit-Debit) End As [CheckSum]					
					
FROM JDT1  T0  with (nolock) LEFT JOIN ORCT T1 with (nolock)  ON T0.BaseRef=T1.DocEntry					
WHERE TransType='24' AND Account IN ('150110-01','150210-01','150220-01','161140-01','223240-01') 
AND T0.RefDate Between @FromDate and @ToDate		
and TrsfrAcct in ('173000-02')	
			
GROUP BY T1.DocNum, T0.ShortName					
,T1.BoeSum,T1.TrsfrSum,T1.CashSum,T1.[CheckSum]

 ) a 					
Group By CardCode
)					
T12 					
ON T0.CardCode=T12.CardCode					
 				
---------------------------------Out Going--------------					
Left Join 					
(					
SELECT OVPM.CardCode,Sum(Case When TrsfrAcct='173000-02' Then TrsfrSum End) As BoeSum,					
                Sum(Case When TrsfrAcct<>'173000-02' Then TrsfrSum End) As TrsfrSum,
                sum(VPM2.SumApplied)-Sum(TrsfrSum) as Adjust					
					
					
 FROM OVPM  with (nolock) Left join OCRD  with (nolock) ON OVPM.CARDCODE=OCRD.CARDCODE 					
           Left Join (Select DocNum,Sum(case when InvType = 13 then -SumApplied else SumApplied end ) as SumApplied FROM VPM2 with (nolock)  Group By DocNuM) VPM2 ON OVPM.DocNum=VPM2.DocNum					
WHERE OCRD.CARDTYPE='C' 
and OVPM.DocDate>=@FromDate AND OVPM.DocDate<=@ToDate 
AND OVPM.Canceled='N' 									
GROUP BY OVPM.CardCode 					
					
)					
  T13 ON T0.CardCode=T13.CardCode					
 					
 					
 Left join   --------------ReconSum					
(					
SELECT   T0.ShortName AS CardCode, SUM( Case WHEN IsCredit='C' Then -ReconSum					
             ELSE ReconSum End) AS ReconSum					
		FROM ITR1 T0			
		--	Left JOIN OINV T2 ON T0.SrcObjAbs=T2.DocEntry AND T0.SrcObjTyp='13'		
--			Left JOIN ORCT T3 ON T0.SrcObjAbs=T3.DocEntry AND T0.SrcObjTyp='24'		
--			Left JOIN ORIN T4 ON T0.SrcObjAbs=T4.DocEntry AND T0.SrcObjTyp='14'		
					
WHERE T0.ReconNum IN (		 			
		 			
		   SELECT T0.BaseRef			
					FROM JDT1 T0  with (nolock) INNER JOIN OJDT T1 with (nolock)  ON T0.TransId=T1.TransId
					WHERE T0.Account IN ('150110-01','150210-01','150220-01','161140-01','223240-01')  AND T1.RefDate BETWEEN @FromDate AND @ToDate
					and  T0.Transtype IN ('321')  )  
					
					AND T0.SrcObjTyp Not IN ('321') And T0.Account in ('150110-01','150210-01','150220-01','161140-01','223240-01')
				--	and T0.ShortName = '131732'
					Group BY T0.ShortName
					
					
) T51					
ON T0.CardCode=T51.CardCode					
 					
 ----------------------------					
					
Left join 					
(SELECT t0.CardCode  ,t1.cardname,SUM(Round(T0.Max1099,2))  as CurInvAmt,SUM(T0.VATSUM) As VatSum					
FROM OINV T0  with (nolock) 					
left join ocrd t1 with (nolock)  on t0.CardCode=t1.cardcode					
where T0.DocDate between @FromDate and @ToDate   and T0.doctype ='I'					
Group by T0.CardCode,t1.cardname) T3					
ON T0.cardcode=T3.CardCode					
 					
Left join 					
(SELECT t0.CardCode,t1.cardname,SUM(Round(T0.Max1099,2)) as CurCrdAmt,SUM(T0.VATSUM) As VatSum					
FROM ORIN T0  with (nolock) 					
left join ocrd t1 with (nolock)  on t0.CardCode=t1.cardcode 	
where 
T0.DocDate between @FromDate and @ToDate	
and T0.DocType='I'				
					
Group by T0.CardCode,t1.cardname  ) T4					
ON T0.cardcode=T4.CardCode					
----------------------					
 ----------------------------					
					
Left join 					
( select  T13.Cardcode  ,   SUM(Freight) as Freight  from 					
(SELECT t0.CardCode  as Cardcode  , CASE when T3.ItmsGrpCod = '139' then    T2.LineTotal  else 0 end as Freight 					
FROM OINV T0  with (nolock) 					
left join INV1 T2 with (nolock)  on T2.DocEntry = T0.DocEntry 					
left join OITM T3 with (nolock)  on T3.ItemCode = T2.ItemCode 					
where T0.DocDate between @FromDate and @ToDate and T0.doctype ='I'					
 					
 union all 					
 					
 SELECT t0.CardCode,  CASE when T3.ItmsGrpCod = '139' then -1 * T2.LineTotal  else 0 end as Freight					
 					
FROM ORIN T0  with (nolock) 					
left join RIN1 T2 with (nolock)  on T2.DocEntry = T0.DocEntry 					
left join OITM T3 with (nolock)  on T3.ItemCode = T2.ItemCode 					
where T0.DocDate between @FromDate and @ToDate and T0.doctype ='I'					
 ) T13 					
 group by t13.Cardcode   ) T14 					
 on  T0.CardCode =T14.Cardcode   					
				
Left Join 					
					
(
SELECT CARDCODE,SUM(ManualJE) as ManualJE FROM(
SELECT ShortName'CARDCODE',SUM(debit-credit) AS ManualJE					
FROM JDT1 with (nolock)  WHERE Account IN ('150110-01','150210-01','150220-01','161140-01','223240-01')  AND RefDate BETWEEN @FromDate AND @ToDate						
and  Transtype='30' 
--AND SHORTNAME='137033'					
Group BY ShortName
union all
select cardcode AS 'CARDCODE',sum(DocTotal)ManualJE
from(
select cardcode,DocTotal	
from OINV a  with (nolock) 
where a.DocType='S' and a.CANCELED='N'
--and CardCode =   '137033' 
AND docDate BETWEEN @FromDate AND @ToDate		
union all
select cardcode 'CARDCODE',-1*DocTotal	
from ORin a  with (nolock) 
where a.DocType='S' and a.CANCELED='N'
--and CardCode =   '137033'  
AND docDate  BETWEEN @FromDate AND @ToDate	
) aa
group by cardcode)EE
group by Cardcode

) T15					
					
ON T0.CardCode=T15.cardcode					
					
where T0.CardType='C' and  t0.CardCode In (SELECT DISTINCT Shortname from JDT1  )  and SlpName='"
                            



$starttime = Get-DATE
$Serv   =  'SZ-SAP01'       
$Database  ='SAPB1_CS'

$SqlConn = New-Object System.Data.SqlClient.SqlConnection
#以 windows 认证连接 MSSQL
$SqlConn.ConnectionString = "Data Source=$Serv;Initial Catalog=$Database;Integrated Security=SSPI;"
 #打开数据库连接
 $SqlConn.open()
  $SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SQL_sales_employee,$SqlConn)
  $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $SqlCmd
  $dataset = New-Object System.Data.DataSet
  $adapter.Fill($dataSet) |out-null  #将sql query的执行结果保存到adapter里面
   
  $qryTable=$dataset.Tables[0]          #[0] this is importandt key ,otherwise ,it will export-csv with system.dataset...system information not correct the value
  $SqlConn.close()
  
  
  Foreach($row in $qryTable)
  {
    $subject_name=""
    $attch_file="c:\temp\CS-Sales&Collection Daily Report_"+$row[1]+ $starttime +".csv"

    $SQL_Report2=$SQL_Report+$row[1]+"'"

    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SQL_Report2,$SqlConn)
    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $SqlCmd
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataSet) |export-csv C:\TEMP\BalanceSheet.csv -NoTypeInformation  #将sql query的执行结果保存到adapter里面
  
  }  
 
# 
 
 #Retrieve  a list of sales employee email which is active and had mail box

  $SqlConn.close()
 #Send mail to each sales employee

 Send-MailMessage -BodyAsHtml -Body $fHTML -From SZ-SZ-SAPB1APP@V.COM  -To Evan.ji@vesuvius.com -SmtpServer "APMailrelay.vesuvius.com" -port 25 -Subject $subject_name -Attachments $attch_file