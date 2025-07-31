<#
  Purpose :  add new user in all sites
  Date    : 2021/7
#>

cls
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'

$ticktNum = "INC0193841 "
$sUsrName ="songemi"

# LDAP Query
Function Fn_Return_AD_Usr_infor ($sUsrName)
{
   $Searcher = New-Object DirectoryServices.DirectorySearcher
        $Searcher.SearchRoot = 'LDAP://DC=corp,DC=vesuvius,DC=com'
        $Searcher.Filter = '(&(objectCategory=user)(cn='+$sUsrName+'))'
        $res = $Searcher.FindAll()  | Sort-Object path
   $Value = "" | Select-Object -Property User_code,first_name ,last_name,email,displayName,phonenum,empid
        
        foreach ($usrTmp in $res)
        {  $Value.User_code = $usrTmp.Properties["name"]
           $Value.first_name = $usrTmp.Properties["sn"] 
            $Value.last_name  = $usrTmp.Properties["givenname"]
            $Value.email = $usrTmp.Properties["mail"] 
            $Value.displayName= $usrTmp.Properties["displayName"]
            $Value.phonenum=$usrTmp.Properties["telephoneNumber"]
            $Value.empid=$usrTmp.Properties["employeeid"]

          #$usrtmp.Properties
        }

  Return $Value
}


 
function fn_AddUser ()  {
     $Val = Fn_Return_AD_Usr_infor $sUsrName
   
     if ($val.User_code -ne $null ) {
         $oUsr=$cmp.getbusinessobject(12)
         $oUsr.UserCode = $Val.User_code[0]
	     $oUsr.UserName = $Val.displayName[0]
	     $oUsr.Email = $Val.email[0]
        # $oUsr.superuser= 1 
	     $oUsr.UserPassword = "Ves-123456"
	     $intRetCode = $oUsr.Add()
 	     if ($intRetCode -eq 0 )
	     { 
		    $oEmployee = $cmp.GetBusinessObject(171)
		    $oEmployee.FirstName =  $Val.first_name[0]
		    $oEmployee.LastName = $Val.last_name[0]
		    $oEmployee.ExternalEmployeeNumber =  $Val.empid[0]
		    $oEmployee.ApplicationUserID =$cmp.GetNewObjectKey()
		    $oEmployee.OfficePhone = $Val.phonenum[0]
		    $oEmployee.Email = $Val.email[0]
		    $oEmployee.Active = 1
		    $outputcode=$oEmployee.Add()
            $outlog =$cmp.CompanyDB + ' '+ $Val.User_code[0] + ' Added Employee with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
          }
	      Else {$outlog =$cmp.CompanyDB + ' '+ $Val.User_code[0] + ' Added  with error code:' +  $intRetCode + ' and error description is:' + $cmp.GetLastErrorDescription()}
      }
      else { $outlog = $sUsrName + ' can not find in Active Directory' }
       return $outlog
}

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
           {  { $site -in "xx"} {$ogeneraldata.SetProperty('U_VES_COmments',$comment)} 
     
              Default { $ogeneraldata.SetProperty('U_VES_Comments',$comment)}
           }
           $ogeneraldata.SetProperty('U_VES_Active',$YesNO)
           $oGeneralServic.Update($ogeneraldata)

           Release-Ref ($oCompServic)
      }
}





$SAP_SiteConnS = @{ 
 AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="CORP\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
 CS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_CS";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
KT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_KT";sapuser="CORP\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
#WE =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WE";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
SQ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_SQ";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}  
WN =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WN";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
#WG =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WG";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
#WV =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WV";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
# HG =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_HG";sapuser="corp\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
}


 $SAP_SiteConnS.Keys|Sort-Object|  ForEach-Object { 
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
       # $cmp.DbUserName=$cmpdbuser
       # $cmp.DbPassword=$cmpdbpwd
        $cmp.UseTrusted=$true
        $cmp.LicenseServer = $cmpLicenseServer

        [void]$cmp.Connect()
       
        if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;Continue
              } 
        else { Write-Host -ForegroundColor Cyan $cmp.CompanyDB connected successfully}

         #disable SAPB1 TN sp control
         Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"

         fn_SAPB1_SP_control $ticktNum 'N' $_

        

          # Add User
         fn_AddUser

         


         #Enable sapb1 TN SP control
         Write-Host -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =''
         fn_SAPB1_SP_control $ticktNum2 'Y' $_
        Write-Host "End processing ."       
       
        $cmp.Disconnect()

  } 

#Release-Ref($cmp)  