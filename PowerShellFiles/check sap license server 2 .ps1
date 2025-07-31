<#
  Purpose: 监控  SAPB1 Compay db server and SLD server 
  Date:20220531

#>
cls
$CurrentProcess = Get-WmiObject Win32_Process | Where ProcessID -match "^$($PID)$" 
If ($CurrentProcess.ProcessName -match '^Powershell.exe$|^Powershell_ise.exe')
{
$CurrentDIR = If ($PSScriptRoot) {$PSScriptRoot} else {Get-Location | Select - expand Path}
} else {
$CurrentDIR = Split-Path (Convert-Path ([environment]::GetCommandLineArgs())[0]) 
} 

$json = Get-Content $CurrentDIR\config.json |ConvertFrom-Json
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
$Date = Get-Date  #job running time
$Server = gc env:computername  #job running server
$ServerList = @{
"SZ-SAPLIC92" =@{db="SLDModel.SLDData";LicenseSvr = "Y"}
"SZ-SAP01" = @{db="SAPB1_WN";dbtype=$json.DbServerType;sapid="Montova"; sappwd ="ButterWN";lic="SZ-SAPLIC92"}
"WG-SAP01" = @{db="SAPB1_WG";dbtype=$json.DbServerType;sapid="Montova"; sappwd ="ButterWG";lic="SZ-SAPLIC92"}
}
$trs = ""

$success = 1 


$HTML='<h1 align="CENTER" style="color: #4485b8;">SAPB1 Server Monitor</h1>
<h4 align="CENTER">$Date</h4>
<table align="CENTER"  style = "border:solid #130c0e; border-width:1px 0px 0px 1px;">
<tbody align="CENTER">
<tr align="CENTER" style="border-top: 2px solid #555;background-color:#d3d7d4">
<td style="border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;">Server    Name </td>
<td style="border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;">DataBase  Name </td>
<td style="border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;">Ping Server</td>
<td style="border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;">MSSQL Service</td>
<td style="border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;">SQL Access Mode</td>
<td style="border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;">Free Space</td>
<td style="border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;">SLD / License   Service</td>
<td style="border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;">Sapb1 Connection</td>
</tr>
$($trs)
</tbody>
</table>'


function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}


 #function to send out mail if meets above anyone 
  function fn_SendMail {
       param ($HtmlBody)
  
         # smtp server
        $emailSmtpServer = "APMailrelay.vesuvius.com"
        $emailSmtpServerPort = "25"
        # recipient 
        $emailFrom = "SAPB1_license@SZ.SZSAPB1APP.com"
        $emailTo = "evan.ji@vesuvius.com"
        # message
        $emailMessage = New-Object System.Net.Mail.MailMessage( $emailFrom , $emailTo )
        $emailMessage.Subject = "$returnStatus connected to  Sapb1 db"
        $emailMessage.IsBodyHtml = $true
        $emailMessage.Priority = $Priority
        $emailMessage.Body = $finalHTML
        #client 
        $SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
        #$SMTPClient.EnableSsl = $True
        #$SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
        $SMTPClient.Send( $emailMessage )  
  }

# 1 ---Test connection to  license server
$ServerList.Keys  | ForEach-Object  {    
    # initial arguments
    $sldSrvic = ""
    $licSrvic = ""
    $NTSrvic = ""
    $DBstates = ""
    $connResult = ""
    $connectedLicenseDscription =""
    $cmpServer = $_
    $cmpCompanyDB = $ServerList[$_]['db']
    $isLic = $ServerList[$_]['LicenseSvr']
	
    IF($isLic -Ne 'Y')
      { 
	    $cmpDbServerType = $ServerList[$_]['dbtype']
        $cmpUserName = $ServerList[$_]['sapid']
        $cmpPassword =$ServerList[$_]['sappwd']
        $licenserServer=$ServerList[$_]['lic']
		$cmp.Server =$cmpServer
	    $cmp.CompanyDB = $cmpCompanyDB
	   $cmp.DbServerType = $cmpDbServerType
	   $cmp.UserName = $cmpUserName
	   $cmp.Password =$cmpPassword
	  $cmp.UseTrusted=$true
      $cmp.SLDServer=$licenserServer
      $connectedLicense = $cmp.Connect()
      $connectedLicenseDscription =$cmp.GetLastErrorDescription()
     $cmp.Disconnect()
	 if ($connectedLicense -eq 0)
	 {  $connectedLicenseDscription = "Successfully Connected"
		$connectedLicense_color="#005831"
		 $connectedLicenseDscription_color="#005831"  # 绿色
		  }
	 else{ $connectedLicense_color="#ed1941"
		   $connectedLicenseDscription_color="#ed1941"
		   $success=0}  # 红色
	}   
	    # ping server 

        $pingSVR=Test-NetConnection -ComputerName $cmpServer
        $ping=$pingSVR.PingSucceeded
        IF ( $ping -EQ $true){ $ping_color="#005831" }
        else   {   $ping_color="#ed1941"  
                 $success= 0		} 
                   
    # check sap server tool service status
  
     Get-CimInstance -ClassName Win32_Service -computername $cmpServer   |Where-Object {$_.Name -eq "B1ServerTools" -or $_.Name -eq  "MSSQLSERVER" -or $_.Name  -eq "B1LicenseService" -or $_.Name -eq "TAO_NT_Naming_Service"}  |
       ForEach-Object {
         IF($isLic -eq 'Y')
          {if ($_.Name -eq "B1ServerTools" )
            {  $b1Serv=$_.State
             if ($b1Serv -eq "Running")         
              {$b1Serv_color="#005831"}
            else{$b1Serv_color="#ed1941"
			$success=0}
            $sldSrvic = "B1ServerTools : " +$b1Serv

               }

          if ($_.Name-eq "TAO_NT_Naming_Service"  )
           {  $TAO_NT_Naming_Service=$_.State
            if ($TAO_NT_Naming_Service -eq "Running")
            { $b1Serv_color="#005831"}
           else{$b1Serv_color="#ed1941"
		   $success=0}
               
            $NTSrvic = "TAO_NT_Naming_Service : "+$TAO_NT_Naming_Service
            }

          if ($_.Name-eq "B1LicenseService"  )
           {  $b1lic=$_.State
            if ($b1lic -eq "Running")
            { $b1Serv_color="#005831"}
           else{$b1Serv_color="#ed1941"
		   $success=0}
             $licSrvic = "B1LicenseService : "+$b1lic
             }
          }   
         if ($_.Name -eq "MSSQLSERVER")
           { $sqlserv=$_.State
             if( $sqlserv -eq "Running" ){
             $sqlserv_color="#005831"}
            else{$sqlserv_color="#ed1941"
			$success=0}
           }
         }
      
          

	# check  db svr  C,D,E Drive free  space
         $DISKS= Get-CIMINSTANCE -Class Win32_LogicalDisk -ComputerName $cmpServer
        FOREACH($DISK IN $DISKS){
         if ($DISK.DeviceID -eq "C:") {
          $diskfreespaceComDB_C= "{0:N2}GB" -F ($DISK.FreeSpace/1GB) 
          if ($diskfreespaceComDB_C -gt 0.2 ){$diskfreespaceComDB_C_color="#005831"}
              else{$diskfreespaceComDB_C_color="#ed1941"
			  $success=0}
			$cmpfreedbC=  "C: "+$diskfreespaceComDB_C
         }
         if ($DISK.DeviceID -eq "D:") {
          $diskfreespaceComDB_D="{0:N2}GB" -F ($DISK.FreeSpace/1GB) 
          if ($diskfreespaceComDB_D -gt 0.2 ){$diskfreespaceComDB_C_color="#005831"}
              else{$diskfreespaceComDB_C_color="#ed1941"
			  $success=0}
			  $cmpfreedbD=  "D: "+$diskfreespaceComDB_D
         }
         if ($DISK.DeviceID -eq "E:") {
          $diskfreespaceComDB_E= "{0:N2}GB" -F ($DISK.FreeSpace/1GB) 
          if ($diskfreespaceComDB_E -gt 0.2 ){$diskfreespaceComDB_C_color="#005831"}
              else{$diskfreespaceComDB_C_color="#ed1941"
			  $success=0}
			  $cmpfreedbE=  "E: "+$diskfreespaceComDB_E
         }
        }
    # 检查 sapb1 db 状态，是否是single user 或者offline ,或者连接数据库出错
	      #创建连接对象
		  $SqlConn = New-Object System.Data.SqlClient.SqlConnection
		 #以 windows 认证连接 MSSQL
		  $SqlConn.ConnectionString = "Data Source=$cmpServer;Initial Catalog=$cmpCompanyDB;Integrated Security=SSPI;"

				
		   #打开数据库连接
		 try { $SqlConn.open()}
		 catch {  $connResult= $_.Exception }

		 $SqlStr = "SELECT    State_Desc, User_Access_Desc  FROM sys.databases where name ='$($cmpCompanyDB)'"
		 $SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlStr,$SqlConn)
		 $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $SqlCmd
		 $dataset = New-Object System.Data.DataSet
		 $adapter.Fill($dataSet) |out-null 
			 
		foreach ($r in  $dataset.Tables[0].Rows)
			  { 
			   $dbstate =  "$($r.State_Desc)"
			   $accessmode ="$($r.User_Access_Desc)"
			  
			  }

		# 显示指定数据库的状态
		if($dbstate -eq 'ONLINE' -and $accessmode -eq 'MULTI_USER' -and !$connResult)
		{   $dbstates_color = "#005831" }
		else 
		{ $dbstates_color = "#ed1941" 
		$success=0}
		$DBstates = “$($dbstate),$($accessmode),connection exception:$connResult"	   
		$sqlconn.Close()

	
		  
   $trs += "<tr align='CENTER' style='background-color: rgba(0,0,0,0.1);'> 
   <td style='border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;font-weight: 400;'>$cmpServer</td>  
      <td style='border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;font-weight: 400;'>$cmpCompanyDB</td>  
<td style='border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;color:$ping_color;font-weight: 400;'>$ping</td>  
<td style='border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;color:$sqlserv_color;font-weight: 400;'>$sqlserv</td>         
<td style='border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;color:$dbstates_color;font-weight: 400;'>$DBstates</td>
<td style='border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;color:$diskfreespaceComDB_C_color;font-weight: 400;'>$cmpfreedbC<br>$cmpfreedbD<br>$cmpfreedbE</td>  
<td style='border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;color:$b1Serv_color;font-weight: 400;'>$sldSrvic<br>$licSrvic<br>$NTSrvic</td>
<td style='border:solid #130c0e; border-width:0px 1px 1px 0px; padding:10px 0px;color:$connectedLicenseDscription_color;font-weight: 400;'>$connectedLicenseDscription</td>
</tr>"
}

# check Citrix publish app service
    #........

    # verify the starttime of sql server and SLD server tools service
     # Get-CimInstance win32_process -computername $licenserServer| 
     #   ? { $_.name -eq "tomcat7.exe" } | 
     #   % { ( $_.CreationDate )}
     if ($success) {$returnStatus ="Successfully";$Priority =1}
     else {$returnStatus="Failure";$Priority=2}
    $finalHTML=$ExecutionContext.InvokeCommand.ExpandString($HTML) 
     fn_SendMail $finalHTML
 

 Release-Ref $cmp


