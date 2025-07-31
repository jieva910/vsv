<#
  purpose: copy local csv file to remote server 
           call SAPB1 DI API version
           update data on the remote server
           sendmail to IT (optional)
#>

$remoteSvr='dg-pol-sapb1app'
$csvfilePath="C:\Temp\ItmCODE.csv"
$Destionation_path = "C:\temp"

 $d=Get-Date -Format "yyyy_MM_ddHHmm"
$log = 'c:\temp\update_Item_puruntWght_'+ $d +'.log' 

#$username = "$server\Administrator"
#$password =  Get-Content C:\mypassword.txt | ConvertTo-SecureString -AsPlainText -Force
#$creds = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password


# copy local csv file to remote server 
$s = New-PSSession -ComputerName $remoteSvr -Credential corp\jievadm -ConfigurationName microsoft.powershell32 

Copy-Item -Path C:\Temp\Itmcode.csv -Destination $Destionation_path  -ToSession $s 

# execute scirpt with csv file on remote server 
Invoke-Command -Session $s -ScriptBlock { 
                    $cmp = New-Object -COMObject 'SAPbobsCOM.Company'
                    function Release-Ref ($ref) {
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers()}
                    function fn_UpdateItem_purchaseweight {
                        param ($Itmcode,$puruntWght)
	
                         $oItmcode=$cmp.getbusinessobject(4)
                        if ( $oItmcode.GetByKey($Itmcode)) {
                            $oItmcode.PurchaseUnitWeight=$puruntWght
                            $outputcode = $oItmcode.Update()
                            $outlog =$cmp.CompanyDB + ' '+ $Itmcode + ' updated  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
                            
                          }
	                    else {$outlog =$Itmcode + ' not exists'} 
                        return $outlog
	 
                    }
                     $starttime =Get-Date
                    $SAP_SiteConnS = @{ SR =@{ Lic="DG-SAPLIC01";db=" DG-POL-SAP01";dbtype="8";cmp="SAPB1_SR";sapuser="jieva";pwd="ves123";DbUserName="Montova";DbPassword="POLmonTova"}
                    }


                     $SAP_SiteConnS.Keys | Sort-Object | ForEach-Object { 
                            $cmpServer = $SAP_SiteConnS[$_]['db']
                            $cmpCompanyDB = $SAP_SiteConnS[$_]['cmp']
                            $cmpDbServerType = $SAP_SiteConnS[$_]['dbtype']
                            $cmpUserName = $SAP_SiteConnS[$_]['sapuser']
                            $cmpPassword =$SAP_SiteConnS[$_]['pwd']
                            $cmpLicenseServer = $SAP_SiteConns[$_]['Lic']
                            $dbusername = $SAP_SiteConns[$_]['DbUserName']
                            $dbpwd= $SAP_SiteConns[$_]['DbPassword']
  
                             $cmp.Server = $cmpServer
                            $cmp.CompanyDB =$cmpCompanyDB
                            $cmp.DbServerType = $cmpDbServerType
                            $cmp.UserName = $cmpUserName
                            $cmp.Password =$cmpPassword
                             $cmp.DbUserName=$dbusername
                             $cmp.DbPassword =$dbpwd
                            #$cmp.UseTrusted=$true
                         
                            $cmp.LicenseServer = $cmpLicenseServer

                            [void]$cmp.Connect()
                            if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() 
                                  break} 

                    # Using local variable ,but this file located on the remote server
                      $CSVfile = Import-Csv  $Using:csvfilePath -Delimiter "," 
    
                     
                    ForEach ($row in $CSVfile){
                               try { 
                                $outlog2=fn_UpdateItem_purchaseweight  $row.itemcode   $row.purchaseweight
                                
			                    }
			                    catch{ $outlog2 = $_.Exception.Message;continue}
	                            $outlog2>> $Using:log 
                 
                             }
                     } 
                      $cmp.Disconnect() | Out-Null 
                       Release-Ref ($cmp)

          # Send-MailMessage -From ps@vsv.com -To 'Evan.ji@vesuvius.com' -SmtpServer 'APMailrelay.vesuvius.com' -Port '25' -Attachments $Log -Subject 'log of udpate' -Body 'check the log please'
  
}  

# copy file from server to local 
Copy-Item -Path $log -Destination $Destionation_path  -FromSession $s 
