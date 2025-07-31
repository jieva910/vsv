<#
  Purpose : Mass update UDT'S udf value in each site
  Date    : 2020/11
#>

cls
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'

$ticktNum = "bank frontend change new ip"

$UserTable = "TIT_SYSFIG"
$new_UDFVal = "10.160.10.201"

function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}

# Control the SP of blocking Super user
Function fn_SAPB1_SP_control($ticknum,$YesNO,$site)
  {   
   $oCompServic = $cmp.GetCompanyService()
   $oGeneralServic = $oCompServic.GetGeneralService('VES_TNMSGS')
   $oGeneralParams = $oGeneralServic.GetDataInterface(3)  # gsGeneralDataParams = 3
   $oGeneralParams.SetProperty('Code','9999999')
   $oGeneralData = $oGeneralServic.GetByParams($oGeneralParams)
   
   # check current SP control status
   $isActive =  $oGeneralData.GetProperty('U_VES_Active')
   $comment = $ticknum
   if ($isActive -ne $YesNO)
     {
          switch ($site)
           {  { $site -in "AS","WG","WV","SZ"} {$ogeneraldata.SetProperty('U_VES_COmments',$comment)} 
     
              Default { $ogeneraldata.SetProperty('U_VES_Comments',$comment)}
           }
           $ogeneraldata.SetProperty('U_VES_Active',$YesNO)
           $oGeneralServic.Update($ogeneraldata)

           Release-Ref ($oCompServic)
      }
}

Function fn_Update_UDf 
{
    $udt = $cmp.GetBusinessObject(153) #Const oUserTables = 153 (&H99)
    $udt = $cmp.UserTables.Item($UserTable)

     IF($udt.GetByKey("A01"))
     {
       $UDT.UserFields.Fields("U_QZIP").value = $new_UDFVal
       Write-host  $cmp.CompanyDB " update UDF value with error code : " $UDT.Update() "and error description: "$cmp.GetLastErrorDescription() 
     }
}


$SAP_SiteConnS = @{  AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
CS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_CS";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
KT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_KT";sapuser="CORP\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
WG =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WG";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
WN =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WN";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
WV =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WV";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
}


 $SAP_SiteConnS.Keys|  ForEach-Object { 
        $cmpServer = $SAP_SiteConnS[$_]['db']
        $cmpCompanyDB = $SAP_SiteConnS[$_]['cmp']
        $cmpDbServerType = $SAP_SiteConnS[$_]['dbtype']
        $cmpUserName = $SAP_SiteConnS[$_]['sapuser']
        $cmpPassword =$SAP_SiteConnS[$_]['pwd']
        $cmpLicenseServer = $SAP_SiteConns[$_]['Lic']
        $cmpdbuser=$SAP_SiteConns[$_]['DbUserName']
        $cmpdbpwd=$SAP_SiteConns[$_]['DbPassword']
  
         $cmp.Server = $cmpServer
        $cmp.CompanyDB =$cmpCompanyDB
        $cmp.DbServerType = $cmpDbServerType
        $cmp.UserName = $cmpUserName
        $cmp.Password =$cmpPassword
        $cmp.DbUserName=$cmpdbuser
        $cmp.DbPassword=$cmpdbpwd
        #$cmp.UseTrusted=$true
        $cmp.LicenseServer = $cmpLicenseServer

        [void]$cmp.Connect()
       
        if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;Continue
              } 
        else { Write-Host -ForegroundColor Green $cmp.CompanyDB connected successfully}

         #disable SAPB1 TN sp control
         Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"

         fn_SAPB1_SP_control $ticktNum 'N' $_

         Write-Host "Start update UDF value in UDT ......"

           fn_Update_UDf


          Write-Host


         #Enable sapb1 TN SP control
         Write-Host -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =''
         fn_SAPB1_SP_control $ticktNum2 'Y' $_
        Write-Host "End processing ."       
       
        $cmp.Disconnect()

  } 

#Release-Ref($cmp)  