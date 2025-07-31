<# Purpose: 
   action: input sitecode for DIAPI connection
           input ticket number for comment inactive SP control
   Date：20200924
#>
$site=read-host "Please Enter SiteCode"  # 1) Enter target site 
$sUsrName=read-host "Please Enter UserCode(5+3)"      # 2 input new usercode 5+3 same as AD account
$sReferencUsrCode = Read-Host "Please enter Reference User code" # 3) input reference to whom 
$ticktNum=Read-Host "Please Enter ticket number"  # 4)  must enter ticket number for live sapb1
$cmp = New-Object -ComObject "sapbobscom.company"
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

  # SAPB1 DI connect to specific site
  Fn_ConnectSAPB1 $cmp $site

  # Disable SP control 
  fn_SAPB1_SP_control $ticktNum 'N'  $site 

 $d=Get-Date -Format "yyyy_MM_ddHHmm"


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


 $starttime =Get-Date


         #disable SAPB1 TN sp control
         Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"
         fn_SAPB1_SP_control $ticktNum 'N'

         Write-Host "Start Adding User ......"
        
         # Add new user in SAPB1
          fn_AddUser 
        
         #Enable sapb1 TN SP control
         #fn_SqlUpdate $sql_activeSP


 Write-Host
 #Enable sapb1 TN SP control
 Write-Host -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
$ticktNum =''
 fn_SAPB1_SP_control $ticktNum 'N'  $site 
Write-Host "End processing ."       
$endtime =Get-Date
$cmp.Disconnect()
Release-Ref($cmp)  
Write-Host -ForegroundColor Red ('running time is :' +($endtime-$starttime).totalseconds)
    