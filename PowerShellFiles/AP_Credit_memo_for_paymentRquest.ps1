# 2024.01.03
 $sql = "SELECT T.DocEntry,T.DocNum,T.DocDate,T.CARDCODE FROM ODPO t INNER JOIN NNM1 n1 ON t.Series = n1.Series
 WHERE LEFT(n1.SeriesName,3) = 'PDI'
 AND t.DocStatus = 'O'"

 
$Serv="SZ-SAPstg91"
$Database ="SAPB1_sz_tst"
$SqlConn = New-Object System.Data.SqlClient.SqlConnection
#连接 MSSQL
   $SqlConn.ConnectionString = "Data Source=$Serv;Initial Catalog=$Database;uid='Butterfly';password='buTterF1y'"

   #打开数据库连接
   $SqlConn.open()
   $SqlCmd = New-Object System.Data.SqlClient.SqlCommand($sql,$SqlConn)
   $SqlCmd.CommandTimeout = 0
  $datareader= $SqlCmd.ExecuteReader() 
  if (!$datareader.HasRows) 
  {$datareader.Close()
  $SqlConn.Close()
  exit}
 

 

$cmp  = new-object -ComObject "sapbobscom.company"

$cmp.Server = $Serv
$cmp.SLDServer ="SZ-TSTSAPLIC92:40000"
$cmp.CompanyDB = $Database
$cmp.DbServerType = 8
$cmp.DbUserName ="Butterfly"
$cmp.DbPassword="buTterF1y"
$cmp.UserName="Montova"
$cmp.Password="ButterSZ"

$rtcode = $cmp.Connect()
IF($rtcode -ne 0)
{Send-MailMessage -Body $cmp.GetLastErrorDescription() -From PayRequest@SAPB1APP.COM  -Subject 'SAPB1 connection failed for payment request' -To 'evan.ji@vesuvius.com'  -SmtpServer 'APMailrelay.vesuvius.com' -port 25
exit }

# AP credit memo
  $oRs_PayRqst = $cmp.GetBusinessObject(300)
  
   $oRs_PayRqst.DoQuery($sql)
   if (!$oRs_PayRqst.EoF)
   {
     [xml]$recordsXML = $oRs_PayRqst.GetAsXML()
      $Nodes = $recordsXML.SelectNodes("//row")
      foreach($N IN $Nodes)
      {
        
        $oPayRequest = $cmp.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oPurchaseDownPayments)
        $oAP = $cmp.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oPurchaseCreditNotes)
 
         [VOID]$oPayRequest.GetByKey($N.DocEntry)

           $linecount = $oPayRequest.Lines.Count
           $OAP.CardCode = $oPayRequest.CardCode
           $OAP.DocDate = Get-Date
           $OAP.DocDueDate = Get-Date

           for($i=0;$i -lt $linecount;$i++)
           {
             $oPayRequest.Lines.SetCurrentLine($i)
               $OAP.Lines.BaseType = $oPayRequest.DocObjectCode
               $OAP.Lines.BaseEntry =  $oPayRequest.DocEntry
               $OAP.Lines.BaseLine = $oPayRequest.Lines.LineNum
                $OAP.Lines.Quantity =  $oPayRequest.Lines.Quantity
        
               $OAP.Lines.add()
           }
   
            $OAP.Comments = " Automatically AP credit memo" 
            [VOID]$OAP.Add()
        Send-MailMessage -Body $OAP.DocNum $cmp.GetLastErrorDescription() -From PayRequest@SAPB1APP.COM  -Subject 'SAPB1 AP credit memo for payment request' -To 'evan.ji@vesuvius.com'  -SmtpServer 'APMailrelay.vesuvius.com' -port 25


      }
   
   
   }
  

