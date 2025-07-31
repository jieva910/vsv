 <#
 purpose : unlock user 
 use     : input sitecode,usercode
 date    :2019/09/29
#>

param($sitecode=$(throw "Parameter missing: -name sitecode") ,
       $usercode=$(throw "Parameter missing: -name usercode") )
"sitecode = $sitecode"
"usercode = $usercode"

function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}

#save site connection in Hash table
$SAP_SiteConn = @{
   AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-1234"}
BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="jieva";pwd="Ves-1234"}
CK =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_CK";sapuser="jieva";pwd="Ves1234"}
CL =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1_CL";sapuser="jieva";pwd="VESint99"}
CN =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_CN";dbtype="6";cmp="SAPB1-CN";sapuser="jieva";pwd="ves123"}
CNBGT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-CN-BGT";sapuser="jieva";pwd="vesint99"}
CS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_CS";sapuser="jieva";pwd="vesint99"}
FC =@{ Lic="DG-SAPLIC01";db="DG-USA-SAP01";dbtype="8";cmp="SAPB1_FC";sapuser="jieva";pwd="ves123"}
FE =@{ Lic="DG-saplic01";db="DG-SAPSQL03\SAPB1_FRA_V9";dbtype="6";cmp="SAPB1-FE";sapuser="jieva";pwd="@Ves123456"}
GH =@{ Lic="DG-SAPLIC01";db="DG-BEL-SAP01";dbtype="8";cmp="SAPB1_GH";sapuser="corp\jieva";pwd="Vesint-999"}
HE =@{ Lic="DG-SAPLIC01";db="DG-BEL-SAP01";dbtype="8";cmp="SAPB1_HE";sapuser="corp\jieva";pwd="Vesint-999"}
HG =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_HG";sapuser="jieva";pwd="Ves-1234"}
IS =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_TUR";dbtype="6";cmp="SAPB1-TUR";sapuser="jieva";pwd="Ves-123456"}
KB =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_KB";sapuser="jieva";pwd="Ves-12345"}
KE =@{ Lic="DG-SAPLIC01";db=" DG-POL-SAP01";dbtype="8";cmp="SAPB1_KE";sapuser="jieva";pwd="Vesint99@"}
KH =@{ Lic="DG-EMEA-LIC01";db="DG-APAC-SAP02\SAPB1_TWN";dbtype="6";cmp="SAPB1-KH";sapuser="jieva";pwd="Ves123"}
KK =@{ Lic="DG-SAPLIC01";db=" DG-POL-SAP01";dbtype="8";cmp="SAPB1_KK";sapuser="jieva";pwd="Vesint99@"}
KX =@{ Lic="DG-SAPLIC01";db=" DG-POL-SAP01";dbtype="8";cmp="SAPB1_KX";sapuser="jieva";pwd="Ves-1234"}
KT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_KT";sapuser="jieva";pwd="vesint99"}
LU =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_LU";dbtype="6";cmp="SAPB1-LU";sapuser="jieva";pwd="ves123"}
MT =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1-MT";sapuser="jieva";pwd="Vesint99"}
MT2 =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1-MT2";sapuser="jieva";pwd="ves123"}
MT3 =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1-MT3";sapuser="jieva";pwd="ves123"}
NC =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_EGY";dbtype="6";cmp="SAPB1-EGY";sapuser="jieva";pwd="VEsint99@"}
OS =@{ Lic="DG-SAPLIC01";db="OS-SAP01";dbtype="8";cmp="SAPB1_OS";sapuser="corp\jieva";pwd="Vesint-999"}
PA =@{ Lic="DG-saplic01";db="DG-SAPSQL03\SAPB1_FRA_V9";dbtype="6";cmp="SAPB1-PA";sapuser="jieva";pwd="VEsint99@"}
PG =@{ Lic="DG-EMEA-LIC01";db="DG-APAC-SAP02\SAPB1_TWN";dbtype="6";cmp="SAPB1-PG";sapuser="jieva";pwd="Ves-123"}
PK =@{ Lic="SZ-SAPLIC8";db="PK-SAPB1";dbtype="6";cmp="VESSBOPK01";sapuser="jieva2";pwd="VEs-123"}
PU =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1-PU";sapuser="corp\jieva";pwd="Vesint-999"}
RK =@{ Lic="DG-EMEA-LIC01";db="RK-SAPB1\SAPB1";dbtype="6";cmp="SAPB1-RK";sapuser="jieva";pwd="Ves123"}
SL =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_SL";sapuser="jieva";pwd="Ves-12345"}
SP =@{ Lic="DG-SAPLIC01";db="DG-ESP-SAP01";dbtype="8";cmp="SAPB1-ESP";sapuser="jieva";pwd="Vesint99@"}
SQ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_SQ";sapuser="jieva";pwd="vesint99"}
SR =@{ Lic="DG-SAPLIC01";db="DG-POL-SAP01";dbtype="8";cmp="SAPB1_SR";sapuser="jieva";pwd="ves123"}
SR2 =@{ Lic="DG-saptst81";db="DG-saptst81";dbtype="8";cmp="SAPB1_SR_tst";sapuser="jieva";pwd="ves123"}
SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="jieva";pwd="Ves-1234"}
TA =@{ Lic="DG-SAPLIC01";db="DG-ITA-SAP01";dbtype="8";cmp="SAPB1_ITA";sapuser="jieva";pwd="Vesint99"}
TC =@{ Lic="DG-SAPLIC01";db="DG-EMEA-SAP01";dbtype="8";cmp="SAPB1_TC";sapuser="jieva";pwd="Ves123"}
TK =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_TK";sapuser="jieva";pwd="Ves1234"}
TW =@{ Lic="DG-EMEA-LIC01";db="DG-APAC-SAP02\SAPB1_TWN";dbtype="6";cmp="SAPB1-TW";sapuser="jieva";pwd="Ves-123"}
VZ =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1_VZ";sapuser="jieva";pwd="VESint99"}
WE =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WE";sapuser="jieva";pwd="vesint99"}
WG =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WG";sapuser="jieva";pwd="vesint99"}
WN =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WN";sapuser="jieva";pwd="vesint99"}
WV =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WV";sapuser="jieva";pwd="vesint99"}


}

 
  
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
$cmp.Server = $SAP_SiteConn.$sitecode['db']
$cmp.CompanyDB = $SAP_SiteConn.$sitecode['cmp']
$cmp.DbServerType = $SAP_SiteConn.$sitecode['dbtype']
$cmp.UserName = $SAP_SiteConn.$sitecode['sapuser']
$cmp.Password =$SAP_SiteConn.$sitecode['pwd']
if ($sitecode -eq 'PG' -OR $sitecode -eq 'KH')
  {$cmp.DbUserName="Butterfly" ;$cmp.DbPassword="buTterF1y"}
elseif ($sitecode -eq'RK') {$cmp.DbUserName="BoomRang";$cmp.DbPassword="B00mrang"}
else  {$cmp.UseTrusted = $True}
$cmp.LicenseServer = $SAP_SiteConn.$sitecode['Lic']

[void]$cmp.Connect()
if($cmp.Connected) {Write-Host $cmp.CompanyName} else {$cmp.GetLastErrorDescription() 
      EXIT}

 
    $usr = $cmp.GetBusinessObject('12')
    $rs = $cmp.GetBusinessObject('300') #recordset

    $rs.doquery(“select userid from ousr where user_code='" +$usercode + "'")
    if ($rs.EoF -eq $false)
     {  $uid = $rs.Fields.Item(0).value
        if ($usr.getbykey($uid) -eq $true ){
       $usr.superuser=0
     $usr.LOCKED=1
         $err_msg = $usr.Update()
           if ($err_msg -eq 0) 
             { Write-Host $usercode  "Remove super role OK"}    
          else{Write-Host $cmp.GetLastErrorDescription()}
       }}
       
  <# else{ $usr.usercode="dahibsag"
     $usr.username="Sagar Dahibhate"
    $usr.superuser=1
    $usr.UserPassword="Ves-123456"
    $usr.add()
    $cmp.GetLastErrorDescription()
           }
           #>
     

  
    <#$usr.usercode=$usercode
    $usr.username='Roman Tomanek'
    $usr.superuser= 1
    $err_msg=$usr.add()#>
   




 
    
    
    #finally{
     #   $processed++
     #   $progresPercent = $processed / $totaltoprocess * 100.0
     #   Write-Progress -Activity "BP Credit limit update in progress" -PercentComplete $progresPercent -Status "$processed out of $totaltoprocess"
    
    #}



$cmp.Disconnect()
Release-Ref($cmp)
 

