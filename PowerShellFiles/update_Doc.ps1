
$site="CS"
$ticktNum="INC0119253"

 $CSVfile = Import-Csv -LiteralPath C:\temp\cs.csv -Delimiter ","


$sql_inactiveSP="UPDATE t SET t.U_VES_Comments ='"+$ticktNum +"',t.U_VES_Active ='N' FROM [@VES_TNMSGS] t WHERE Code = '9999999'"
$sql_activeSP="UPDATE t SET t.U_VES_Comments ='"+$ticktNum +"',t.U_VES_Active ='Y' FROM [@VES_TNMSGS] t WHERE Code = '9999999'"
 $d=Get-Date -Format "yyyy_MM_ddHHmm"
 $log = 'c:\temp\' + $site + '_inativebp_' + $d +'.log'
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'

function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}
function fn_UpdateDoc {
    param ($docentry)
      $po = $cmp.GetBusinessObject(22)
    if ( $po.GetByKey($docentry) ) {
        $po.UserFields.Fields("U_Ves_Toinv").value ='Y'
      
        #$po.FrozenFrom = get-date -Format "yyyy-MM-dd"
        #$po.FrozenTo = "2999-12-31"
        
        $outputcode = $po.Update()
        $outlog =$cmp.CompanyDB + ' '+ $docentry + ' updated  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
        
      }
       else {$outlog = $docentry + ' not exists'} 
       return $outlog
}
function fn_SqlUpdate {
     param ($sqlstr)
     $oRs = $cmp.GetBusinessObject('300') #recordset
     $oRs.doquery($sqlstr)
}
 $starttime =Get-Date
$SAP_SiteConnS = @{  AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
CK =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_CK";sapuser="jieva";pwd="Ves1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
CL =@{ Lic="DG-SAPLIC01";db="DG-IND-SAP02";dbtype="8";cmp="SAPB1_CL";sapuser="jieva";pwd="VESint99";DbUserName="Montova";DbPassword="CLmonTova"}
CN =@{ Lic="DG-EMEA-LIC01";db="DG-SAPSQL01\SAPB1_CN";dbtype="6";cmp="SAPB1-CN";sapuser="jieva";pwd="ves123";DbUserName="monTova";DbPassword="ButterSA"}
CNBGT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-CN-BGT";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
CS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_CS";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
CSTST =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPtst82";dbtype="8";cmp="SAPB1_CS_tst";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
FC =@{ Lic="DG-SAPLIC01";db="DG-USA-SAP01";dbtype="8";cmp="SAPB1_FC";sapuser="jieva";pwd="ves123";DbUserName="Butterfly";DbPassword="buTterF1y"}
FE =@{ Lic="DG-SAPLIC01";db="DG-SAPSQL03\SAPB1_FRA_V9";dbtype="6";cmp="SAPB1-FE";sapuser="jieva";pwd="Ves-123";DbUserName="monTova";DbPassword="PAmonTova"}
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
        $cmp.DbUserName=$cmpdbuser
        $cmp.DbPassword=$cmpdbpwd
        #$cmp.UseTrusted=$true
        $cmp.LicenseServer = $cmpLicenseServer

        $cmp.Connect()
        if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() 
              BREAK} 
  } 
         Write-Host "Start processing ......"
         #disable SAPB1 TN sp control
         #fn_SqlUpdate $sql_inactiveSP
         
         #update property of  docentrys in CSV file
         ForEach ($row in $CSVfile){           
                    try { $outlog2=fn_UpdateDoc  $row.docentry }
                    catch { $outlog2= $_.Exception.Message;continue }
                  write-host  $outlog2
                 }
         #Enable sapb1 TN SP control
        # fn_SqlUpdate $sql_activeSP

 
Write-Host "End processing ."       
$endtime =Get-Date
$cmp.Disconnect()
Release-Ref($cmp)  
Write-Host -ForegroundColor Red ('running time is :' +($endtime-$starttime).totalseconds)
    