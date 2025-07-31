<# Purpose: copy data between company ,like quick copy function
   action: input sitecode for DIAPI connection
           input ticket number for comment inactive SP control
   Date：20200929
   CSV File format like below:
   UserTables	UserFields	UserObjects	QueryCategories	UserQueries	StoredProcedures	FactoryIndicator
    VES_CONSGNDTL	VES_ItemCode	VES_CONSGN	[SZ] Stock	Consignment Stock Posting History	VES_TN_L_GRPOConsignment	VC
    VES_CONSGNDTL	VES_ItemName		VES] FMS	VES_Get Consignment ItemCode	VES_TN_L_ConsignmentStock	
    VES_CONSGNDTL	VES_UOM		VES] FMS	VES_Get Consignment ItemName		
    VES_CONSGNDTL	VES_Qty		VES] FMS	VES_Get Consignment ItemUom		
    VES_CONSGNDTL	VES_Reason		VES] FMS	VES_Get Consignment SupplierCode		
    VES_CONSGNDTL	VES_Remarks		VES] FMS	VES_Get Consignment SupplierName		
    VES_CONSGNHDR	VES_CardCode		VES] FMS	VES_GetDate Consignment Stock		
    VES_CONSGNHDR	VES_CardName					
    VES_CONSGNHDR	VES_Date					
    VES_CONSGNHDR	VES_Ref					
#>

$site="SZ"             #choose source site 


<# using csv instead of below 
$UserTables = @("VES_CONSGNDTL","VES_CONSGNHDR")
$UserFields =@{"@VES_CONSGNDTL"=@("VES_ItemCode","VES_ItemName","VES_UOM","VES_Qty","VES_Reason","VES_Remarks")
               "@VES_CONSGNHDR"=@("VES_CardCode","VES_CardName","VES_Date","VES_Ref") }
$UserObjects=@("VES_CONSGN")
$UserQueries=@{"[SZ] Stock"=@("Consignment Stock Posting History")
               "[VES] FMS"=@() }
$QueryCategories=@("[SZ] Stock","[VES] FMS")
$StoredProcedures=@()
#>
$DataList = "C:\Temp\Copy2companyDB\UDT_UDO_Udf_query_SP_list.csv"
$savePath ="c:\Temp\Copy2companyDB\"

$ticktNum="INC0119340"
$sql_inactiveSP="UPDATE t SET t.U_VES_Comments ='"+$ticktNum +"',t.U_VES_Active ='N' FROM [@VES_TNMSGS] t WHERE Code = '9999999'"
$sql_activeSP="UPDATE t SET t.U_VES_Comments ='"+$ticktNum +"',t.U_VES_Active ='Y' FROM [@VES_TNMSGS] t WHERE Code = '9999999'"
 $d=Get-Date -Format "yyyy_MM_ddHHmm"
 $log = 'c:\temp\' + $Target_site + '_CopyUDF_UDO_UDT_SP_' + $d +'.log'



function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}


function fn_SqlUpdate {
     param ($sqlstr)
     $oRs = $cmp.GetBusinessObject('300') #recordset
     $oRs.doquery($sqlstr)
}

$starttime =Get-Date

$SAP_SiteConnS = @{  SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
}

$CSVfile = Import-Csv -LiteralPath $DataList -Delimiter ","

# 1 ---connect to source sapb1 company db--------------------------------

$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
$SAP_SiteConnS.Keys -match $site |ForEach-Object { 
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
        [void]$cmp.Connect()       
        if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() 
              } 
         else {Write-Host $cmp.CompanyDB Connected successfully}
}
         Write-Host
         Write-Host "Start Export process ......"
    
         
# 2 ----------------------------------- Export object ---------------------------------------
         $cmp.XmlExportType = 3 
         $factoryindicator = $cmp.GetBusinessObject(138)   #oFactoringIndicators = 138 (&H8A)
         $udo = $cmp.GetBusinessObject(206)                 #Const oUserObjectsMD = 206 (&HCE)
         $oUQ = $cmp.GetBusinessObject(160) #Const oUserQueries = 160 (&HA0)
         $oUF = $cmp.GetBusinessObject(152) #Const oUserFields = 152 (&H98)
         $oUT = $cmp.GetBusinessObject(153) #Const oUserTables = 153 (&H99)
          $rs = $cmp.GetBusinessObject('300') #recordset
          $querysql="SELECT t.IntrnalKey,t.QCategory FROM OUQR t WHERE t.QName ='"
          $SPsql="SELECT text FROM sys.syscomments t WHERE object_name(id)='"
     
     Foreach( $obj in $CSVfile)
     {
       If ($factoryindicator.GetByKey($obj.FactoryIndicator)){$factoryindicator.SaveXML($savepath + $obj.FactoryIndicator + ".xml")}     
       If ($oUT.GetByKey($obj.UserTables) ){$oUT.SaveXML($savepath + $obj.UserTables + ".xml")}
       If ($udo.GetByKey($obj.UserObjects) ){$udo.SaveXML($savepath + $obj.UserObjects + ".xml")}
       
       # Export user query
       $rs.DoQuery($querysql+$obj.UserQueries + "'")
         if (!$rs.EoF)
          {
           $queryID=$rs.Fields.Item(0).value
           $sourceCategroyID=$rs.Fields.Item(1).value
           If ($oUQ.GetByKey($queryID,$sourceCategroyID)){$oUQ.SaveXML($savepath + $obj.UserQueries + ".xml")}
          }  

        # Export stored procedure to .sql file
        if ($obj.StoredProcedures -ne $null)
         {$rs.DoQuery($SPsql + $obj.StoredProcedures +"'")
          if (!$rs.EoF)
            { 
             
             $SPfile =  $savepath+$obj.StoredProcedures +'.sql'
             set-Content   -path $SPfile -Value $rs.Fields.Item(0).value -ErrorAction SilentlyContinue
            }
          } 
     }

     # Export UDFs
     $Temps=$CSVfile |Group-Object UserTables | Sort-Object UserTables | Select-Object name,count
     foreach($temp in $Temps)
     {
      for($i=0;$i -lt $temp.count;$i++) 
      {
        If ($oUF.GetByKey('@'+$temp.Name,$i) ){$oUF.SaveXML($savepath + $temp.Name + $i.ToString() + ".xml")}
      }
     
     }
      Write-Host
      Write-Host "End exporting process ." 
   #  $cmp.Disconnect()
  # ---------------------------------------------------------------------------------------------------------
         
   



# 3 --- connect to Target sapb1 company db and Import objects--------------------------------
$cmp2 = New-Object -COMObject 'SAPbobsCOM.Company'
$Target_site = "cstst"  #choose target site 

$DataList2 = "C:\Temp\Copy2companyDB\UDT_UDO_Udf_query_SP_list.csv"
$savePath2 ="c:\Temp\Copy2companyDB\"
$CSVfile2 = Import-Csv -LiteralPath $DataList2 -Delimiter ","


$ticktNum="INC0119340"
$sql_inactiveSP="UPDATE t SET t.U_VES_Comments ='"+$ticktNum +"',t.U_VES_Active ='N' FROM [@VES_TNMSGS] t WHERE Code = '9999999'"
$sql_activeSP="UPDATE t SET t.U_VES_Comments ='"+$ticktNum +"',t.U_VES_Active ='Y' FROM [@VES_TNMSGS] t WHERE Code = '9999999'"

$SAP_TargetSites = @{  AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
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
 $SAP_TargetSites.Keys -match $Target_site |  ForEach-Object { 
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
       
        if(-not $cmp2.Connected) {Write-Host $cmp2.GetLastErrorDescription() 
              } 
        else { Write-Host $cmp2.CompanyDB connected successfully}
}
  Write-Host "Start processing ......"
#disable SAPB1 TN sp control
#fn_SqlUpdate $sql_inactiveSP

#-------------Execute Import Objects -------------------------
         $cmp2.XmlExportType = 3 
         $factoryindicator2 = $cmp2.GetBusinessObject(138)   #oFactoringIndicators = 138 (&H8A)
         $udo2 = $cmp2.GetBusinessObject(206)                 #Const oUserObjectsMD = 206 (&HCE)
         $oUQ2 = $cmp2.GetBusinessObject(160) #Const oUserQueries = 160 (&HA0)
         
         $oUT2 = $cmp2.GetBusinessObject(153) #Const oUserTables = 153 (&H99)
         $rs2 = $cmp2.GetBusinessObject('300') #recordset

         #$ReturnCodeCollection = ""| Select-Object -Property Action,Error,Errdiscription

          # Import User Tables
           $UserTables=$CSVfile2|Group-Object UserTables | Select-Object name
             foreach ($UT IN $UserTables)
            {
                 $oUT2=$cmp2.GetBusinessObjectFromXML( $savepath2 + $UT.name + ".xml",0)
                  Write-Host  "Add user Table "  $UT.name "with error code: "   $oUT2.Add()  " and error description: "  $cmp2.GetLastErrorDescription()
          
            }
           Release-Ref ($oUT2)


          # Import UDFs
         $UDFs=$CSVfile2 |Group-Object UserTables | Sort-Object UserTables | Select-Object name,count
         foreach($UDF in $UDFs)
         {
          for($i=0;$i -lt $UDF.count;$i++) 
          {  $oUF2 = $cmp2.GetBusinessObjectFromXML($savepath2 + $UDF.Name + $i.ToString() + ".xml",0) #Const oUserFields = 152 (&H98)
           write-host "Add UDFs " $UDF.Name "with error code: " $oUF2.add() " and error description: " $cmp2.GetLastErrorDescription()
          }
         }

         Release-Ref ($oUF2)

         $querysql2="SELECT t.CategoryId FROM OQCN t WHERE t.CatName like '%"
         foreach($obj2 in $CSVfile2)
         { #Import UDO
           If ($obj2.UserObjects -ne $null) {$udo2.Browser.ReadXml($savepath2 + $obj2.UserObjects + ".xml",0)
              $ReturnCodeCollection.Action="Add UDO $obj2.UserObjects"
             $ReturnCodeCollection.Error=$udo2.add()
             $ReturnCodeCollection.Errdiscription=$cmp2.GetLastErrorDescription()
             }
              # Import User Queries
           if ($obj2.QueryCategories -ne $null) 
             {  $cateName = $obj2.QueryCategories -split ' '
               $rs2.DoQuery($querysql2+ $cateName[1]+ "'")
               if (!$rs.EoF)
                {  #retrieve category id of target site  query manager
                   $target_CategoryID=$rs2.Fields.Item(0).value
                  [xml]$xmlUser_query=Get-Content -Path $savepath2 + $obj2.UserQueries + ".xml"
                  $xmlUser_query.BOM.bo.UserQueries.row.QueryCategory=$target_CategoryID
                 $xmlUser_query.Save($savepath2 + $obj2.UserQueries + ".xml") 
                  $oUQ2.Browser.ReadXml($savepath +$obj2.UserQueries+ ".xml",0)
                  $ReturnCodeCollection.Action="Add User query $obj2.UserQueries"
                  $ReturnCodeCollection.Error=$oUQ2.add()
                  $ReturnCodeCollection.Errdiscription=$cmp2.GetLastErrorDescription()
              }  
             }
          }
 Write-Host
Write-Host "End processing ."  
$ReturnCodeCollection | Out-File c:\temp\returncollection.log     
$endtime =Get-Date
#Enable sapb1 TN SP control
#fn_SqlUpdate $sql_activeSP
#$cmp2.Disconnect()
#Release-Ref($cmp)  
#Write-Host -ForegroundColor Red ('running time is :' +($endtime-$starttime).totalseconds)
    