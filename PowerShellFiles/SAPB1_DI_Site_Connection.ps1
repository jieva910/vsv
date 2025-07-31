<#
  purpose : this is library of SAP di api company connection of all sites
  Date    : 2020.11
#>

# All sites SAP company connection information list 
 # store SAPB1 company db in Hash table


$SAP_SiteConnS = @{  AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
    BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
    yktst =@{ Lic="SZ-tstSAPLIC92";db="SZ-SAPSTG91";dbtype="8";cmp="SAPB1_YK_Tst";sapuser="manager";pwd="ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
 YK =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_YK";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
    yktst3 =@{ Lic="SZ-tstSAPLIC92";db="SZ-SAPSTG91";dbtype="8";cmp="SAPB1_YK_Tst3";sapuser="manager";pwd="ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}

    CK =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_CK";sapuser="jieva";pwd="Ves1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
    CKstg =@{ Lic="DG-SAPTST81";db="DG-SAPSTG91";dbtype="8";cmp="SAPB1_CK_STG";sapuser="\";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
    CL =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1_CL";sapuser="jieva";pwd="VESint99";DbUserName="Montova";DbPassword="CLmonTova"}
    CN =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_CN";dbtype="6";cmp="SAPB1-CN";sapuser="jieva";pwd="ves123";DbUserName="monTova";DbPassword="ButterSA"}
    budTST =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPSTG91";dbtype="8";cmp="SAPB1-CN-BGT_TST";sapuser="jinman";pwd="ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
    CNBGT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-CN-BGT";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
    CS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_CS";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
    CSTST =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPSTG91";dbtype="8";cmp="SAPB1_CS_TST";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
     CSTST2 =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPSTG91";dbtype="8";cmp="SAPB1_CS_tst2";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
    BYTST =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPSTG91";dbtype="8";cmp="SAPB1_by_tst2";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
    FC =@{ Lic="DG-SAPLIC01";db="DG-USA-SAP01";dbtype="8";cmp="SAPB1_FC";sapuser="jieva";pwd="ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
    FE =@{ Lic="DG-SAPLIC01";db="DG-SAPSQL03\SAPB1_FRA_V9";dbtype="6";cmp="SAPB1-FE";sapuser="jieva";pwd="Ves-123";DbUserName="monTova";DbPassword="PAmonTova"}
    fetst =@{ Lic="DG-SAPTST81";db="DG-TSTSAPSQL02\SAPB1_3";dbtype="6";cmp="SAPB1_FE_ToBe";sapuser="jieva";pwd="ves123";DbUserName="monTova";DbPassword="PAmonTova"}
    HG =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_HG";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
    IS =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_TUR";dbtype="6";cmp="SAPB1-TUR";sapuser="jieva";pwd="Ves-123456";DbUserName="monTova";DbPassword="TRmonTova"}
    KB =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_KB";sapuser="jieva";pwd="Ves-12345";DbUserName="Butterfly";DbPassword="buTterF1y"}
    KE =@{ Lic="DG-SAPLIC01";db="DG-POL-SAP01";dbtype="8";cmp="SAPB1_KE";sapuser="jieva";pwd="Vesint99@";DbUserName="Montova";DbPassword="POLmonTova"}
    KH =@{ Lic="DG-EMEA-LIC01";db="DG-APAC-SAP02\SAPB1_TWN";dbtype="6";cmp="SAPB1-KH";sapuser="jieva";pwd="Ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
    KK =@{ Lic="DG-SAPLIC01";db="DG-POL-SAP01";dbtype="8";cmp="SAPB1_KK";sapuser="jieva";pwd="Vesint99@";DbUserName="Montova";DbPassword="POLmonTova"}
    KX =@{ Lic="DG-SAPLIC01";db="DG-POL-SAP01";dbtype="8";cmp="SAPB1_KX";sapuser="jieva";pwd="Ves-1234";DbUserName="Montova";DbPassword="POLmonTova"}
    KT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_KT";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
    KTtst =@{ Lic="sz-TSTSAPLIC92";db="SZ-SAPSTG91";dbtype="8";cmp="SAPB1_KT_TST";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
    LU =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_LU";dbtype="6";cmp="SAPB1-LU";sapuser="jieva";pwd="ves123"}
    MT =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1-MT";sapuser="jieva";pwd=",NOVesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
    MT2 =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1-MT2";sapuser="jieva";pwd="ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
    MT3 =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1-MT3";sapuser="jieva";pwd="ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
    NC =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_EGY";dbtype="6";cmp="SAPB1-EGY";sapuser="jieva";pwd="Ves-123456";DbUserName="Butterfly";DbPassword="buTterF1y"}
    PA =@{ Lic="DG-SAPLIC01";db="DG-SAPSQL03\SAPB1_FRA_V9";dbtype="6";cmp="SAPB1-PA";sapuser="\";pwd="Vesint-999";DbUserName="monTova";DbPassword="PAmonTova"}
    PG =@{ Lic="DG-EMEA-LIC01";db="DG-APAC-SAP02\SAPB1_TWN";dbtype="6";cmp="SAPB1-PG";sapuser="jieva";pwd="Ves-123";DbUserName="Butterfly";DbPassword="buTterF1y"}
    PK =@{ Lic="PK-SAPB1";db="PK-SAPB1";dbtype="6";cmp="VESSBOPK01";sapuser="jieva";pwd="Ves-123456";DbUserName="montova";DbPassword="PKmonTova"}
    PKtst =@{ Lic="PK-SAPB1";db="PK-SAPB1";dbtype="6";cmp="SAPB1-PK-STAGING";sapuser="jieva";pwd="Ves-123456";DbUserName="montova";DbPassword="PKmonTova"}
    PU =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1-PU";sapuser="jieva";pwd="VEsint99@";DbUserName="montova";DbPassword="CLmonTova"}
    RK =@{ Lic="DG-EMEA-LIC01";db="RK-SAPB1\SAPB1";dbtype="6";cmp="SAPB1-RK";sapuser="jieva";pwd="Vesu@12345";DbUserName="BoomRang";DbPassword="B00mrang"}
    SL =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_SL";sapuser="jieva";pwd="Ves-12345";DbUserName="Butterfly";DbPassword="buTterF1y"}
    LA =@{ Lic="DG-SAPLIC01";db="DG-ESP-SAP01";dbtype="8";cmp="SAPB1-ESP";sapuser="jieva";pwd="Vesint99@";DbUserName="monTova";DbPassword="LAmonTova"}
    SQ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_SQ";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
    SR =@{ Lic="DG-SAPLIC01";db="DG-POL-SAP01";dbtype="8";cmp="SAPB1_SR";sapuser="jieva";pwd="ves123";DbUserName="Montova";DbPassword="POLmonTova"}
    SRTST =@{ Lic="DG-SAPTST81";db="DG-POL-SAPSTG91";dbtype="8";cmp="SAPB1_SR_STG";sapuser="jieva";pwd="ves123";DbUserName="Montova";DbPassword="POLmonTova"}
    SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="jieva";pwd="Ves-123456";DbUserName="Butterfly";DbPassword="buTterF1y"}
    SZTST =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPSTG91";dbtype="8";cmp="SAPB1_SZ_tst";sapuser="jinman";pwd="Ves-123456";DbUserName="Butterfly";DbPassword="buTterF1y"}
    GE =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1_ITA";sapuser="jieva";pwd="Vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
    TC =@{ Lic="DG-SAPLIC01";db="DG-EMEA-SAP01";dbtype="8";cmp="SAPB1_TC";sapuser="jieva";pwd="Ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
    TK =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_TK";sapuser="jieva";pwd="Ves1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
    TW =@{ Lic="DG-EMEA-LIC01";db="DG-APAC-SAP02\SAPB1_TWN";dbtype="6";cmp="SAPB1-TW";sapuser="jieva";pwd="Ves-123"}
    VZ =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1_VZ";sapuser="jieva";pwd="VESint99";DbUserName="Montova";DbPassword="CLmonTova"}
    WE =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WE";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
    WG =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WG";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
    WN =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WN";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
    WV =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WV";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
    OS =@{ Lic="DG-SAPLIC01";db="OS-SAP01";dbtype="8";cmp="SAPB1_OS";sapuser="\";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
    GH =@{ Lic="DG-SAPLIC01";db="DG-BEL-SAP01";dbtype="8";cmp="SAPB1_GH";sapuser="\";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
    HE =@{ Lic="DG-SAPLIC01";db="DG-BEL-SAP01";dbtype="8";cmp="SAPB1_HE";sapuser="\";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
    LY =@{ Lic="DG-SAPLIC01";db="DG-BEL-SAP01";dbtype="8";cmp="SAPB1_LY";sapuser="\";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
   
    HN =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1_HN";sapuser="\";pwd="Vesint-999";DbUserName="montova";DbPassword="CLmonTova"}
    VF =@{ Lic="DG-SAPLIC01";db="DG-EMEA-SAP01";dbtype="8";cmp="SAPB1_VF";sapuser="jieva";pwd="Vesu@12345";DbUserName="Butterfly";DbPassword="buTterF1y"}
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