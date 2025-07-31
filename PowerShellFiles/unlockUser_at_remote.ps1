$remoteSvr = "dg-jpn-sapb1app"
$Target_site ='ck'
$usercode = 'kurodkei'
$s = New-PSSession -ComputerName $remoteSvr -Credential corp\jievadm -ConfigurationName microsoft.powershell32 

Invoke-Command -Session $s -ScriptBlock { 
$cmp2 = New-Object -COMObject 'SAPbobsCOM.Company'
$SAP_TargetSites = @{  AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
AStst =@{ Lic="sz-TSTSAPLIC92";db="SZ-SAPTST82";dbtype="8";cmp="SAPB1_as_tst";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
BYtst =@{ Lic="sz-TSTSAPLIC92";db="SZ-SAPTST82";dbtype="8";cmp="SAPB1_BY_tst";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
CK =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_CK";sapuser="jieva";pwd="Ves1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
CL =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1_CL";sapuser="jieva";pwd="VESint99";DbUserName="Montova";DbPassword="CLmonTova"}
CN =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_CN";dbtype="6";cmp="SAPB1-CN";sapuser="jieva";pwd="ves123";DbUserName="monTova";DbPassword="ButterSA"}
CNBGT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-CN-BGT";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
CS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_CS";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
CSTST =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPtst82";dbtype="8";cmp="SAPB1_CS_tst";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
FC =@{ Lic="DG-SAPLIC01";db="DG-USA-SAP01";dbtype="8";cmp="SAPB1_FC";sapuser="jieva";pwd="ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
FE =@{ Lic="DG-SAPLIC01";db="DG-SAPSQL03\SAPB1_FRA_V9";dbtype="6";cmp="SAPB1-FE";sapuser="corp\jieva";pwd="Vesint-999";DbUserName="monTova";DbPassword="PAmonTova"}
fetst =@{ Lic="DG-SAPTST81";db="DG-TSTSAPSQL02\SAPB1_3";dbtype="6";cmp="SAPB1_FE_ToBe";sapuser="jieva";pwd="ves123";DbUserName="monTova";DbPassword="PAmonTova"}
HG =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_HG";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
IS =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_TUR";dbtype="6";cmp="SAPB1-TUR";sapuser="jieva";pwd="Ves-123456";DbUserName="monTova";DbPassword="TRmonTova"}
KB =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_KB";sapuser="jieva";pwd="Ves-12345";DbUserName="Butterfly";DbPassword="buTterF1y"}
KE =@{ Lic="DG-SAPLIC01";db=" DG-POL-SAP01";dbtype="8";cmp="SAPB1_KE";sapuser="jieva";pwd="Vesint99@";DbUserName="Montova";DbPassword="POLmonTova"}
KH =@{ Lic="DG-EMEA-LIC01";db="DG-APAC-SAP02\SAPB1_TWN";dbtype="6";cmp="SAPB1-KH";sapuser="jieva";pwd="Ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
KK =@{ Lic="DG-SAPLIC01";db=" DG-POL-SAP01";dbtype="8";cmp="SAPB1_KK";sapuser="jieva";pwd="Vesint99@";DbUserName="Montova";DbPassword="POLmonTova"}
KX =@{ Lic="DG-SAPLIC01";db=" DG-POL-SAP01";dbtype="8";cmp="SAPB1_KX";sapuser="jieva";pwd="Ves-1234";DbUserName="Montova";DbPassword="POLmonTova"}
KT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_KT";sapuser="CORP\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
KTtst =@{ Lic="sz-TSTSAPLIC92";db="SZ-SAPTST82";dbtype="8";cmp="SAPB1_KT_tst";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
LU =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_LU";dbtype="6";cmp="SAPB1-LU";sapuser="jieva";pwd="ves123"}
MT =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1-MT";sapuser="jieva";pwd=",NOVesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
MT2 =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1-MT2";sapuser="jieva";pwd="ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
MT3 =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1-MT3";sapuser="jieva";pwd="ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
NC =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_EGY";dbtype="6";cmp="SAPB1-EGY";sapuser="jieva";pwd="VEsint99@";DbUserName="Butterfly";DbPassword="buTterF1y"}
PA =@{ Lic="DG-SAPLIC01";db="DG-SAPSQL03\SAPB1_FRA_V9";dbtype="6";cmp="SAPB1-PA";sapuser="corp\jieva";pwd="Vesint-999";DbUserName="monTova";DbPassword="PAmonTova"}
PG =@{ Lic="DG-EMEA-LIC01";db="DG-APAC-SAP02\SAPB1_TWN";dbtype="6";cmp="SAPB1-PG";sapuser="jieva";pwd="Ves-123";DbUserName="Butterfly";DbPassword="buTterF1y"}
PK =@{ Lic="PK-SAPB1";db="PK-SAPB1";dbtype="6";cmp="VESSBOPK01";sapuser="jieva";pwd="Ves-123456";DbUserName="montova";DbPassword="PKmonTova"}
PU =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1-PU";sapuser="jieva";pwd="VEsint99@";DbUserName="montova";DbPassword="CLmonTova"}
RK =@{ Lic="DG-EMEA-LIC01";db="RK-SAPB1\SAPB1";dbtype="6";cmp="SAPB1-RK";sapuser="jieva";pwd="Ves123";DbUserName="BoomRang";DbPassword="B00mrang"}
SL =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_SL";sapuser="jieva";pwd="Ves-12345";DbUserName="Butterfly";DbPassword="buTterF1y"}
LA =@{ Lic="DG-SAPLIC01";db="DG-ESP-SAP01";dbtype="8";cmp="SAPB1-ESP";sapuser="jieva";pwd="Vesint99@";DbUserName="monTova";DbPassword="LAmonTova"}
SQ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_SQ";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
SR =@{ Lic="DG-SAPLIC01";db=" DG-POL-SAP01";dbtype="8";cmp="SAPB1_SR";sapuser="jieva";pwd="ves123";DbUserName="Montova";DbPassword="POLmonTova"}
SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
SZTST =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPtst82";dbtype="8";cmp="SAPB1_SZ_tst";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
GE =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1_ITA";sapuser="jieva";pwd="Vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
TC =@{ Lic="DG-SAPLIC01";db="DG-EMEA-SAP01";dbtype="8";cmp="SAPB1_TC";sapuser="jieva";pwd="Ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
TK =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_TK";sapuser="jieva";pwd="Ves1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
TW =@{ Lic="DG-EMEA-LIC01";db="DG-APAC-SAP02\SAPB1_TWN";dbtype="6";cmp="SAPB1-TW";sapuser="jieva";pwd="Ves-123"}
VZ =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1_VZ";sapuser="jieva";pwd="VESint99";DbUserName="Montova";DbPassword="CLmonTova"}
WE =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WE";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
WG =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WG";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
WN =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WN";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
WV =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WV";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
OS =@{ Lic="DG-SAPLIC01";db="OS-SAP01";dbtype="8";cmp="SAPB1_OS";sapuser="corp\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
GH =@{ Lic="DG-SAPLIC01";db="DG-BEL-SAP01";dbtype="8";cmp="SAPB1_GH";sapuser="corp\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
HE =@{ Lic="DG-SAPLIC01";db="DG-BEL-SAP01";dbtype="8";cmp="SAPB1_HE";sapuser="corp\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
HN =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1_HN";sapuser="corp\jieva";pwd="Vesint-999";DbUserName="montova";DbPassword="CLmonTova"}
}
  # must full match sitecode ^ $
 $SAP_TargetSites.Keys -match "^"+$Using:Target_site+"$"  |  ForEach-Object { 
        $cmpServer2 = $SAP_TargetSites[$_]['db']
        $cmpCompanyDB2 = $SAP_TargetSites[$_]['cmp']
        $cmpDbServerType2 = $SAP_TargetSites[$_]['dbtype']
        $cmpUserName2 = $SAP_TargetSites[$_]['sapuser']
        $cmpPassword2 =$SAP_TargetSites[$_]['pwd']
        $cmpLicenseServer2 = $SAP_TargetSites[$_]['Lic']
        $cmpdbuser2=$SAP_TargetSites[$_]['DbUserName']
        $cmpdbpwd2=$SAP_TargetSites[$_]['DbPassword']
  
        $cmp2.Server = $cmpServer2
        $cmp2.CompanyDB =$cmpCompanyDB2
        $cmp2.DbServerType = $cmpDbServerType2
        $cmp2.UserName = $cmpUserName2
        $cmp2.Password =$cmpPassword2
        $cmp2.DbUserName=$cmpdbuser2
        $cmp2.DbPassword=$cmpdbpwd2
        #$cmp.UseTrusted=$true
        $cmp2.LicenseServer = $cmpLicenseServer2

        [void]$cmp2.Connect()
       
        if(-not $cmp2.Connected) {Write-Host $cmp2.GetLastErrorDescription() ;break
              } 
        else { Write-Host -ForegroundColor Cyan $cmp2.CompanyDB connected successfully}
}

$usr = $cmp2.GetBusinessObject('12')
    $rs = $cmp2.GetBusinessObject('300') #recordset

    $rs.doquery(“select userid from ousr where user_code='" +$Using:usercode + "'")
    if ($rs.EoF -eq $false)
     {  $uid = $rs.Fields.Item(0).value
        if ($usr.getbykey($uid) -eq $true )
        {$usr.locked =0 ;
          # $usr.UserPassword ='Ves-123456'
          #$usr.SUPERUSER=1
            $err_msg = $usr.Update()
           if ($err_msg -eq 0) 
            { Write-Host $Using:usercode  Unlock OK}
           else{Write-Host $cmp2.GetLastErrorDescription()}
         }
     }
      else{Write-Host $Using:usercode  not exiting}


}