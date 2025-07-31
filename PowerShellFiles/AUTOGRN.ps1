cls

$dbsvr ="dg-ai-global-01"
$db ="BusinessIntegration"
$sql_DBconnections = "SELECT t.UserName,t.UserPassword,t.DatabaseName
  ,t.DBUserName,t.DBPassword,t.ServerName,t.LicenceServer FROM SiteConnection t WHERE t.SiteCode IN ('AS','kt','BY','cs','sz','wn','YK') 
  AND isnull(t.AppName,'') =''"
 # get DocEntry list of  open PO ,AUTO GRN table existing in AS live ,SZ TST DB.
 $sql_DocEntry  = "
  SELECT distinct p.cardcode,p.docnum,p1.DocEntry,p1.linenum,p1.opencreqty quantity,CAST(p10.LineText AS NVARCHAR(256)) SpecialLineText FROM OPOR p INNER JOIN POR1 p1 ON p.DocEntry = p1.DocEntry
  INNER JOIN [sapb1_YK].dbo.[@ves_autogrn] ag ON p1.ItemCode = ag.u_itemcode 
 -- AND P1.Dscription = ag.U_Description
  AND p1.U_Ves_Requisitioner = ag.u_requestor
  and p.CardName = ag.U_suppliername
  LEFT JOIN POR10 p10 ON p10.DocEntry = p1.DocEntry  AND p10.ObjType = p1.ObjType AND p10.AftLineNum = p1.LineNum AND p10.LineType = 'T'
  inner join oadm a on   a.manager = ag.u_site
  LEFT JOIN ORDR t2 ON t2.NumAtCard = RIGHT(CAST(P10.LineText AS NVARCHAR(200)),10)
  WHERE p1.LineStatus = 'o' AND p.Printed = 'Y' and ISNULL(t2.docstatus,'C')='C'  and p.DocDate>='20240407'
"

FUNCTION HasRows($sapdbsvr,$sapdb,$sapdbuser,$sapdbpwd,$sqlquery)
{
  $sqlconn_SAPDB   =New-Object System.Data.SqlClient.SqlConnection
  $sqlconn_SAPDB.ConnectionString = "Data Source=$sapdbsvr;Initial Catalog=$sapdb;UID=$sapdbuser;password=$sapdbpwd"
  $sqlconn_SAPDB.Open()
  $SqlCmd_SAPDB = New-Object System.Data.SqlClient.SqlCommand
  $SqlCmd_SAPDB.connection = $SqlConn_SAPDB
  $SqlCmd_SAPDB.commandtext = $sqlquery
   $dbreader  = $SqlCmd_SAPDB.ExecuteReader() 
   if (!$dbreader.HasRows){$rt = 0} 
   else {$rt = 1 }
   return $rt
   $dbreader.close()
   $sqlconn_SAPDB.close()
}

$SqlConn = New-Object System.Data.SqlClient.SqlConnection
$SqlConn.ConnectionString = "Data Source=$dbsvr;Initial Catalog=$db;Integrated Security=SSPI;"
$SqlConn.open()
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.connection = $SqlConn
$SqlCmd.commandtext = $sql_DBconnections
$tbs  = $SqlCmd.ExecuteReader() 
while ($tbs.Read())
{
  $hasVal = HasRows $tbs.GetValue(5) $tbs.GetValue(2) $tbs.GetValue(3) $tbs.GetValue(4) $sql_DocEntry
  if ($hasVal)
 { $cmp = New-Object -ComObject 'sapbobscom.company'
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
       $oRs = $cmp.getbusinessobject(300)
      $oRs.doquery($sql_DocEntry)
      $xml =[xml]$oRs.getasxml()
      $nodes = $xml.selectnodes("//row")
      $groupbyDocEntry = $nodes|  Group-Object -Property Cardcode,DocEntry

     foreach($grps in $groupbyDocEntry)
      {  
        $oGRPO = $cmp.GetBusinessObject(20);
	     $oGRPO.CardCode = ($grps.Name -split ",")[0]  ;
	     $oGRPO.DocDate = get-date -Format "yyyy-MM-dd"  ;
        foreach($grp in $grps.Group)
         {  
           $i = 0 
           $oGRPO.lines.basetype = 22 ;
	       $oGRPO.lines.baseentry = $grp.DocEntry ;
	       $oGRPO.lines.baseline =[INT]$grp.linenum   ;
	       $oGRPO.lines.Quantity =$grp.quantity  ;
	       $oGRPO.Lines.SetCurrentLine($i) ;
	       $oGRPO.lines.add() ;
           $oGRPO.SpecialLines.LineType =[SAPbobsCOM.BoDocSpecialLineType]::dslt_Text
           $oGRPO.SpecialLines.AfterLineNumber = [INT]$oGRPO.lines.LineNum-1     # note this value
           $oGRPO.SpecialLines.LineText =$grp.SpecialLineText
           # $oGRPO.SpecialLines.SetCurrentLine($k)
            $oGRPO.SpecialLines.Add()
           $i++
         }
      $rtCode = $oGRPO.add()
    if ($rtCode -eq 0 ) {$msg = "$($grp.docnum) has been finished Auto GRN in $($cmp.CompanyDB)" }
    else {$msg = "$($grp.docnum) failed to auto GRN in $($cmp.CompanyDB) ,$($cmp.GetLastErrorDescription())"}
     Send-MailMessage -Body $msg -From AutoGRN@SAPB1APP.COM  -Subject 'SAPB1 Auto GRN' -To 'evan.ji@vesuvius.com'  -SmtpServer 'APMailrelay.vesuvius.com' -port 25
     
 }
      [void]$cmp.Disconnect()
   }
  else {$msg = "failed to connect to $($cmp.CompanyDB),auto try later ,$($cmp.GetLastErrorDescription())" 
    
   Send-MailMessage -Body $msg -From AutoGRN@SAPB1APP.COM  -Subject 'SAPB1 Auto GRN' -To 'evan.ji@vesuvius.com'  -SmtpServer 'APMailrelay.vesuvius.com' -port 25
    break
  }
   
}

}
  
$tbs.Close()
$SqlConn.close()



 