cls

$dbsvr ="dg-ai-global-91"   # "dg-ai-global-01"
$db = "AI_TEST"              # "BusinessIntegration"
                            # 'as','kt','by','cs','wn'
$sql_DBconnections = "SELECT t.UserName,t.UserPassword,t.DatabaseName
  ,t.DBUserName,t.DBPassword,t.ServerName,t.LicenceServer FROM SiteConnection t WHERE t.SiteCode IN ('sz') 
  AND isnull(t.AppName,'') =''"

 # $sql_StoredPrc  = "exec Storedprocedure"
 $sql_DocEntry  = "select * from dbo.VES_GRNS where Status='O'"

 
$SqlConn = New-Object System.Data.SqlClient.SqlConnection
$SqlConn.ConnectionString = "Data Source=$dbsvr;Initial Catalog=$db;Integrated Security=SSPI;"
$SqlConn.open()
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.connection = $SqlConn
$SqlCmd.commandtext = $sql_DBconnections
$tbs  = $SqlCmd.ExecuteReader() 
while ($tbs.Read())
{
  $cmp = New-Object -ComObject 'sapbobscom.company'
  $cmp.Server=$tbs.GetValue(5)
  $cmp.CompanyDB=$tbs.GetValue(2)
  $cmp.LicenseServer=$tbs.GetValue(6)
  $cmp.DbUserName= $tbs.GetValue(3)
  $cmp.DbPassword=$tbs.GetValue(4)
  $cmp.DbServerType = 8
  $cmp.UserName= $tbs.GetValue(0)
  $cmp.Password= $tbs.GetValue(1)
  
  $rtCode = $cmp.Connect()
  if ($rtCode -eq 0 ) {
       $oRs_sp = $cmp.getbusinessobject(300)
       $oRs = $cmp.getbusinessobject(300)
      #  $oRs_sp.doquery($sql_StoredPrc)
      $oRs.doquery($sql_DocEntry)
      if ($oRs.Eof){break}
      $xml =[xml]$oRs.getasxml()
      $nodes = $xml.selectnodes("//row")
      $groupbyDocEntry = $nodes|  Group-Object -Property Cardcode,SDNumber,SDInvDate

     foreach($grps in $groupbyDocEntry)
      {  
        $oOPCH = $cmp.GetBusinessObject(18);
	     $oOPCH.CardCode = ($grps.Name -split ",")[0]  ;
	     $oOPCH.DocDate = get-date -Format "yyyy-MM-dd"  ;
          $dateInt = [int](($grps.Name -split ",")[2])  ;
         $oOPCH.TaxDate = "{0:####-##-##}" -f  $dateInt 
        foreach($grp in $grps.Group)
         {  
           $i = 0 
           $oOPCH.lines.basetype = 20 ;
	       $oOPCH.lines.baseentry = $grp.GRNEntry ;
	       $oOPCH.lines.baseline =$grp.linenum   ;
	       $oOPCH.lines.Quantity =$grp.GRNOpenQty  ;
	       $oOPCH.Lines.SetCurrentLine($i) ;
	       $oOPCH.lines.add() ;
           $i++
         }
      $rtCode = $oOPCH.add()
    if ($rtCode -eq 0 ) {$msg = "has been finished Auto APINVOICE in $($cmp.CompanyDB)" }
    else {$msg = "failed to auto APINVOICE in $($cmp.CompanyDB) ,$($cmp.GetLastErrorDescription())"}
     Send-MailMessage -Body $msg -From AutoAPInvoice@SAPB1APP.COM  -Subject 'SAPB1 AutoAPInvoice' -To 'evan.ji@vesuvius.com'  -SmtpServer 'APMailrelay.vesuvius.com' -port 25
     
 }
      [void]$cmp.Disconnect()
   }
  else {$msg = "failed to connect to $($cmp.CompanyDB),auto try later ,$($cmp.GetLastErrorDescription())" 
    
   Send-MailMessage -Body $msg -From AutoAPInvoice@SAPB1APP.COM  -Subject 'SAPB1 AutoAPInvoice' -To 'evan.ji@vesuvius.com'  -SmtpServer 'APMailrelay.vesuvius.com' -port 25
    break
  }
   
}

  
$tbs.Close()
$SqlConn.close()

