cls
$CurrentProcess = Get-WmiObject Win32_Process | Where ProcessID -match "^$($PID)$" 
If ($CurrentProcess.ProcessName -match '^Powershell.exe$|^Powershell_ise.exe')
{
$CurrentDIR = If ($PSScriptRoot) {$PSScriptRoot} else {Get-Location | Select - expand Path}
} else {
$CurrentDIR = Split-Path (Convert-Path ([environment]::GetCommandLineArgs())[0]) 
} 
$json = Get-Content $CurrentDIR\config.json |ConvertFrom-Json
$dbsvr =$json.AIServer
$db =$json.AIDB
$sql_DBconnections = $json.AIConnectionstring
 # get DocEntry list of  open PO ,AUTO GRN table existing in AS live ,SZ TST DB.
 $sql_DocEntry  = $json.sql_autoGRN

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
  $cmp.DbServerType =$json.DbServerType
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



 