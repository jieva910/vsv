cls
 
<#  dg-ai-global-01 
SELECT 
      t.[ReportId]
	  ,t1.ReportName
      ,[LegalEntityCode]
	  ,t1.DocDate
	 -- ,'2024-12-30' 'SAPPostDate'
      ,t1.TaxDate 
	  ,t2.ExpenseType
	  ,t2.AccountCode
	  ,t2.PostedAmount
	  ,t2.CostCenter
	  ,t2.DIM2
	  ,t2.DIM3
	  ,t2.DIM4
	  ,t2.DIM5
	  ,t2.Remarks 'Comments'
      
  FROM [ApplicationIntegration_Data].[dbo].[ConcurReportDetail] t 
  inner join [ApplicationIntegration_Data].[dbo].[ConcurReportHeader] t1 on t.ReportId = t1.ReportId
 inner join [ApplicationIntegration_Data].[dbo].[ConcurReportLine] t2 on t1.ReportId = t2.ReportId
  where t.CountryCode ='cn'  and t.CreatedDate >'20241228' and t1.DocDate <='20241231' and IsProcessed = 0


#>


$dbsvr ="dg-ai-global-91"
$db ="AI_TEST"
$sql_DBconnections = "SELECT t.UserName,t.UserPassword,t.DatabaseName
  ,t.DBUserName,t.DBPassword,t.ServerName,t.LicenceServer FROM SiteConnection t WHERE t.SiteCode IN ('sz') 
  AND isnull(t.AppName,'') =''"
 


$SqlConn = New-Object System.Data.SqlClient.SqlConnection
$SqlConn.ConnectionString = "Data Source=$dbsvr;Initial Catalog=$db;Integrated Security=SSPI;"
$SqlConn.open()
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.connection = $SqlConn
$SqlCmd.commandtext = $sql_DBconnections
$tbs  = $SqlCmd.ExecuteReader() 
while ($tbs.Read())
{
  $hasVal =1
  if ($hasVal)
 { $cmp = New-Object -ComObject 'sapbobscom.company'
  $cmp.Server=$tbs.GetValue(5)
  $cmp.CompanyDB=$tbs.GetValue(2)
  $cmp.LicenseServer=$tbs.GetValue(6)
  $cmp.DbUserName= $tbs.GetValue(3)
  $cmp.DbPassword=$tbs.GetValue(4)
  $cmp.DbServerType = 8    #sql server 2014
  $cmp.UserName= $tbs.GetValue(0)
  $cmp.Password= $tbs.GetValue(1)
  
  $rtCode = $cmp.Connect()
  if ($rtCode -eq 0 ) {
      $csv = Import-Csv 'C:\Dell\SZ pending.csv'
		$groupbyDocEntry = $csv| Group-Object  DIM3,ReportID,TaxDate,ReportName
     foreach($grps in $groupbyDocEntry)
      {  
        $oAP = $cmp.GetBusinessObject(18);
	     $oAP.CardCode = ($grps.Name -split ",")[0]  ;
		   $oAP.NumAtCard= ($grps.Name -split ",")[1]  ;
	    # $oAP.DocDate = get-date -Format "yyyy-MM-dd"  ;
		 $oAP.DocDate = '2024-12-31' ;
		  $oAP.TaxDate = ($grps.Name -split ",")[2]  ;
		  $oAP.Comments = ($grps.Name -split ",")[3]  ;
		  $oAP.DocType = 1 
        foreach($grp in $grps.Group)
         {  
           $i = 0  
		    $oAP.lines.ItemDescription=$grp.ExpenseType
           $oAP.lines.AccountCode = $grp.AccountCode ;
              $oAP.lines.UnitPrice = [DOUBLE]$grp.PostedAmount;
	       $oAP.lines.CostingCode =$grp.CostCenter ;
	       $oAP.lines.CostingCode2 = $grp.DIM2 ;
	       $oAP.lines.CostingCode3 =$grp.DIM3 ;
		   $oAP.lines.CostingCode4 = $grp.DIM4;
		   $oAP.lines.CostingCode5 =$grp.DIM5 ;
		 
	       $oAP.Lines.SetCurrentLine($i) ;
	       $oAP.lines.add() ;
           $i++
         }
      $rtCode = $oAP.add()
    if ($rtCode -eq 0 ) {$msg = "$(($grps.Name -split ",")[1]) has been finished  in $($cmp.CompanyDB)" }
    else {$msg = "$(($grps.Name -split ",")[1]) failed  in $($cmp.CompanyDB) ,$($cmp.GetLastErrorDescription())"}
	Write-Host $msg
    # Send-MailMessage -Body $msg -From AutoGRN@SAPB1APP.COM  -Subject 'SAPB1 Auto GRN' -To 'evan.ji@vesuvius.com'  -SmtpServer 'APMailrelay.vesuvius.com' -port 25
     
 }
      [void]$cmp.Disconnect()
   }
  else {Write-Host  "failed to connect to $($cmp.CompanyDB),auto try later ,$($cmp.GetLastErrorDescription())" 
    
 #  Send-MailMessage -Body $msg -From AutoGRN@SAPB1APP.COM  -Subject 'SAPB1 Auto GRN' -To 'evan.ji@vesuvius.com'  -SmtpServer 'APMailrelay.vesuvius.com' -port 25
    break
  }
   
}

}
  
$tbs.Close()
$SqlConn.close()



 