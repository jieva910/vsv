
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}
function fn_UpdateBP {
    param ($bpcode,$website,$idno,$pwd)
	
     $oBP=$cmp.getbusinessobject(2)
    if ( $oBP.GetByKey($bpcode)) {
        $oBP.website=$website
         $oBP.AdditionalID=$idno
        $oBP.Password=$pwd
                $outputcode = $oBP.Update()
        $outlog =$cmp.CompanyDB + ' '+ $bpcode + ' updated  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
        return $outlog
      }
	 
	 
}
 $starttime =Get-Date
$SAP_SiteConnS = @{ TK =@{ Lic="DG-SAPTST81";db="DG-SAPSTG91";dbtype="8";cmp="SAPB1_CK_STG";sapuser="jieva";pwd="Ves1234";DbUserName="Butterfly";DbPassword="buTterF1y"}}


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
         $cmp.usetrusted=$true
        if ($_ -eq 'PG' -OR $_ -eq 'KH')
          {$cmp.DbUserName="Butterfly" ;$cmp.DbPassword="buTterF1y"}
        elseif ($_ -eq'RK') {$cmp.DbUserName="BoomRang";$cmp.DbPassword="B00mrang"}
        else  {$cmp.UseTrusted = $True}
        $cmp.LicenseServer = $cmpLicenseServer

        [void]$cmp.Connect()
        if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() 
              EXIT} 


  $CSVfile = Import-Csv -LiteralPath C:\Temp\BPCODE.csv -Delimiter "," 
    
  $d=Get-Date -Format "yyyy_MM_ddHHmm"
  $log = 'c:\temp\update_bp_website_'+ $d +'.log' 

 ForEach ($row in $CSVfile){
           try { 
             $outlog2=fn_UpdateBP  $row.BPCode   $row.'Web Site' $row.'ID No. 2' $row.Password
			}
			catch{ $outlog2 = $_.Exception.Message;continue}
	        $outlog2>> $log 
                 
         }
  } 
  $cmp.Disconnect() | Out-Null 
   Release-Ref ($cmp)

  # Send-MailMessage -From ps@vsv.com -To 'Evan.ji@vesuvius.com' -SmtpServer 'APMailrelay.vesuvius.com' -Port '25' -Attachments $Log -Subject 'log of udpate' -Body 'check the log please'
