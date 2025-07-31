<#
  purpose : this is library of SAP di api company connection of all sites
  Date    : 2020.11
#>

# All sites SAP company connection information list 
 # store SAPB1 company db in Hash table


$SAP_SiteConnS = @{  
    VF =@{ Lic="DG-SAPLIC01";db="DG-EMEA-SAP01";dbtype="8";cmp="SAPB1_VF";sapuser="jieva";pwd="";DbUserName="";DbPassword=""}
    }

Function Fn_ConnectSAPB1 ($cmp,$site)
{   
   
    # connect to specific site 
  $SAP_SiteConnS.Keys -match "^"+$site+"$" | Sort-Object | ForEach-Object { 
        $cmpServer = $SAP_SiteConnS[$site]['db']
        $cmpCompanyDB = $SAP_SiteConnS[$site]['cmp']
        $cmpDbServerType = $SAP_SiteConnS[$site]['dbtype']
        $cmpUserName = $SAP_SiteConnS[$site]['sapuser']
        $cmpPassword =$SAP_SiteConnS[$site]['pwd']
        $cmpLicenseServer = $SAP_SiteConns[$site]['Lic']
        $cmpdbuser=$SAP_SiteConns[$site]['DbUserName']
        $cmpdbpwd=$SAP_SiteConns[$site]['DbPassword']
  
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
       
        if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;break
              } 
        else { Write-Host -ForegroundColor Cyan $cmp.CompanyDB connected successfully}
  }
}


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
           {  { $site -in "XX"} {$ogeneraldata.SetProperty('U_VES_COmments',$comment)} 
     
              Default { $ogeneraldata.SetProperty('U_VES_Comments',$comment)}
           }
           $ogeneraldata.SetProperty('U_VES_Active',$YesNO)
           $oGeneralServic.Update($ogeneraldata)

           Release-Ref ($oCompServic)
      }
}




# Release COM object 
function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}



############################################ RunSpacePool -Parall ##########################################
#---------------enable InitialSessionState-----
$InitialSessionState=[System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
#-------------get all function in current session---
 Get-ChildItem Function:\ | Where-Object {$_.Name -notlike "*:*"}|select name -ExpandProperty name|
 ForEach-Object{$definition=Get-Content "function:\$_" -ErrorAction Stop
 $function=New-object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $_ ,$definition
 $InitialSessionState.Commands.Add($function) }

#---------------enable Runspacepool-----
$Runspacepool=[runspacefactory]::createRunspacepool($InitialSessionState)
[void]$Runspacepool.SetMinRunspaces(1)
[void]$Runspacepool.SetMaxRunspaces(5)
$Runspacepool.Open()
$powershell=[powershell]::create()
$powershell.Runspacepool=$Runspacepool
$hash=[hashtable]::Synchronized(@{})
$jobs = New-Object System.Collections.ArrayList

###########################################  RunSpacePool -Parall ##########################################

## BP Master Data

# 1
Function fn_Update_salesEmployee($BP,$salesEmpID)
{
  $salesEmp = $cmp.GetBusinessObject(2)
  if ($salesEmp.GetBykey($BP))
  {
    $salesEmp.SalesPersonCode = $salesEmpID
    Write-Host $BP 'update sales employee error: ' $salesEmp.update() $cmp.GetLastErrorDescription()
  }
}

# 2
Function fn_update_BPcreditlimit($BP,$creditlmt,$maxcommit)
{
  $oBP=$cmp.getbusinessobject(2)
    if ( $oBP.GetByKey($BP.Trim()) ) {
        $oBP.CreditLimit =$creditlmt
        $oBP.MaxCommitment = $maxcommit
        $outputcode = $oBP.Update()
        $outlog =$cmp.CompanyDB + ' '+ $BP + ' updated  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
         Write-Host
      }
}
