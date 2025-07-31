
$cmp2 = New-Object -COMObject 'SAPbobsCOM.Company'

$Target_site = read-host "Please Enter SiteCode"  # 1) Enter target site 

$savepath2= 'c:\Temp\Copy2companyDB\'
$DataList2 = "C:\Temp\Copy2companyDB\UDT_UDO_Udf_query_SP_list.csv"

$CSVfile2 = Import-Csv -LiteralPath $DataList2 -Delimiter ","

$ticktNum=Read-Host "Please Enter ticket number"  # 2)  must enter ticket number for live sapb1

#realease DI objects
 function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}


Function fn_SAPB1_SP_control
  {  param ( $ticknum ,$YesNO)
   $oCompServic = $cmp2.GetCompanyService()
   $oGeneralServic = $oCompServic.GetGeneralService('VES_TNMSGS')
   $oGeneralParams = $oGeneralServic.GetDataInterface(3)  # gsGeneralDataParams = 3
   $oGeneralParams.SetProperty('Code','9999999')
   $oGeneralData = $oGeneralServic.GetByParams($oGeneralParams)
   $ogeneraldata.SetProperty('U_VES_Comments',$ticktNum)
   $ogeneraldata.SetProperty('U_VES_Active',$YesNO)
   $oGeneralServic.Update($ogeneraldata)

   Release-Ref ($oCompServic)
  }

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
  # must full match sitecode ^ $
 $SAP_TargetSites.Keys -match "^"+$Target_site+"$"  |  ForEach-Object { 
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
       
        if(-not $cmp2.Connected) {Write-Host $cmp2.GetLastErrorDescription() ;exit
              } 
        else { Write-Host -ForegroundColor Cyan $cmp2.CompanyDB connected successfully}
}

#disable SAPB1 TN sp control
 Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"
 fn_SAPB1_SP_control $ticktNum 'N'


  Write-Host  -ForegroundColor Green "Start importing User tables  ......"
   $oUT2 = $cmp2.GetBusinessObject(153) #Const oUserTables = 153 (&H99)
    $cmp2.XmlExportType = 3 
   
    # A.import User tables
     foreach ($UT IN $CSVfile2.UserTables)
    {    if($UT){
         $oUT2=$cmp2.GetBusinessObjectFromXML( $savepath2 + $UT + ".xml",0)
         Write-Host  "Add user Table "  $UT "with error code: "   $oUT2.Add()  " and error description: "  $cmp2.GetLastErrorDescription()
         }
    }
       Release-Ref ($oUT2)
    Write-Host  -ForegroundColor Green "End importing User tables  ......"
    Write-Host
    Write-Host

    Write-Host  -ForegroundColor Green "Start importing User fields  ......"
      # Import UDFs
         $UDFs=$CSVfile2 |Group-Object UserTables | Sort-Object UserTables | Select-Object name,count
         foreach($UDF in $UDFs)
         {
          for($i=0;$i -lt $UDF.count;$i++) 
          { $oUF2 = $cmp2.GetBusinessObjectFromXML($savepath2 + $UDF.Name + $i.ToString() + ".xml",0) #Const oUserFields = 152 (&H98)
           write-host "Add UDFs "  $savepath2$UDF.Name  $i.ToString()  ".xml with error code: " $oUF2.add() " and error description: " $cmp2.GetLastErrorDescription()
          }
         }
         Release-Ref ($oUF2)
     Write-Host  -ForegroundColor Green "End importing User fields  ......"
    Write-Host
    Write-Host 
    
       
       # B.import UDO
       Write-Host  -ForegroundColor Green "Start importing UDO ......"
        foreach($UDO IN $CSVfile2.UserObjects)
        { if  ($UDO) {
          $xml_udo_file=$savepath2 + $UDO + ".xml"
         $udo2=$cmp2.GetBusinessObjectFromXML($xml_udo_file,0)
        Write-Host "Add UDO " $xml_udo_file " with error code: " $udo2.add() " and error description: " $cmp2.GetLastErrorDescription()
         }
        }

         Release-Ref ($udo2)
         Write-Host  -ForegroundColor Green "End importing UDO......"
         Write-Host
         Write-Host   

          # C. Import User Queries
           Write-Host  -ForegroundColor Green "Start importing User Queries ......"
           $rs2 = $cmp2.GetBusinessObject('300') #recordset
            $querysql2="SELECT t.CategoryId FROM OQCN t WHERE t.CatName = '"
           foreach($obj2 in $CSVfile2)
           {if ($obj2.SourceQueryCategories) 
             {  $cateName = $obj2.SourceQueryCategories -split ' '
                if($cateName[1] -eq "FMS"){$Target_cateName = $obj2.SourceQueryCategories}else{$Target_cateName ="["+$Target_site.substring(0,2).ToUpper()+"] "+ $cateName[1]  }
               $rs2.DoQuery($querysql2+ $Target_cateName+ "'")
               if (!$rs2.EoF){$target_CategoryID=$rs2.Fields.Item(0).value}
                 #retrieve category id of target site  query manager
                   $xmlfile= $savepath2 + $obj2.UserQueries + ".xml"
                  [xml]$xmlUser_query=Get-Content -Path $xmlfile
                  $xmlUser_query.BOM.bo.UserQueries.row.QueryCategory=$target_CategoryID.ToString()
                 $xmlUser_query.Save($xmlfile) 
                  $oUQ2=$cmp2.GetBusinessObjectFromXML($xmlfile,0)
                Write-Host  "Add User query" $obj2.UserQueries "with error code:" $oUQ2.add() " error description:" $cmp2.GetLastErrorDescription()  }
             
            }

            Release-Ref ($rs2)
            Write-Host  -ForegroundColor Green "End importing User Queries......"
            Write-Host
            Write-Host   

            # D. import Transaction Notification message (UDO data)
              Write-Host  -ForegroundColor Green "Start importing Transaction Notification message ......"
                $UDO_DATA = Import-Csv C:\Temp\Copy2companyDB\UDO_Data.csv
                $oCMPService = $cmp2.GetCompanyService()
                $oGenearlService=$oCMPService.GetGeneralService('VES_TNMSGS')   # UDO table
           
                foreach($cell in $UDO_DATA){
                  $oGenearlData = $oGenearlService.GetDataInterface(1) # 1 = gsGeneralData
                  $oGenearlData.SetProperty("Code",$cell.Code)
                  $oGenearlData.SetProperty("Name",$cell.Name)
                   $oGenearlData.SetProperty("U_VES_FrgnName",$cell.U_VES_FrgnName) 	
                    $oGenearlData.SetProperty("U_VES_Name",$cell.U_VES_Name) 
              
                  [void]$oGenearlService.Add($oGenearlData)
                   $cmp2.GetLastErrorDescription()
                }

                 Release-Ref ($oCMPService)
               Write-Host  -ForegroundColor Green "End importing Transaction Notification message......"
               Write-Host
               Write-Host 

          # E . Import sql function in sapb1 db
            Write-Host  -ForegroundColor Green "Start importing  sql function ......"
             $sql_function="create FUNCTION [dbo].[fn_VES_ConsignmentStock](
                @CardCode VARCHAR(20),
                @ItemCode VARCHAR(20),
                @Date Datetime
                )
                RETURNS @Table Table (CardCode1 VARCHAR(20),
                ItemCode1 VARCHAR(20),
                Stock1 float, 
                Date1 Datetime) AS
                BEGIN
                Insert Into @Table (CardCode1, ItemCode1, Stock1, Date1)
                Select K.BpCode, K.ItemCode ,SUM(K.Stock), MAX(K.TransactionDate) As Stock FROM
                (SELECT X.BpCode, X.ItemCode, X.Stock,  X.FromDate AS 'TransactionDate' FROM (
                SELECT T0.U_VES_CardCode AS BpCode, T1.U_VES_ItemCode AS ItemCode, T1.U_VES_Qty AS Stock, T0.U_Ves_Date AS FromDate
                FROM [dbo].[@VES_CONSGNHDR] T0  INNER JOIN [dbo].[@VES_CONSGNDTL] T1 ON T0.[DocEntry] = T1.[DocEntry] 
                Union All
                SELECT T0.CardCode, T1.ItemCode, -T1.Quantity, T0.DocDate Type  FROM OPDN T0 Inner Join PDN1 T1 ON T0.DocEntry=T1.DocEntry
                Where T0.Indicator='VC'
                Union All
                SELECT T0.CardCode, T1.ItemCode, T1.Quantity, T0.DocDate FROM ORPD T0 Inner Join RPD1 T1 ON T0.DocEntry=T1.DocEntry
                Where T0.Indicator='VC'
                ) X) K
                WHERE K.TransactionDate<=@Date AND K.BpCode=@CardCode and K.ItemCode=@ItemCode
                Group By K.BpCode, K.ItemCode
                Return
                END;"
              $oRs_function = $cmp2.GetBusinessObject(300)
               $oRs_function.DoQuery($sql_function) 
                Release-Ref ($oRs_function)
                 Write-Host  -ForegroundColor Green "End importing  sql function......"
               Write-Host
               Write-Host 


          # F. import stored procedures to live sapb1 db

          Write-Host  -ForegroundColor Green "Start importing  stored procedures SQL ......"
            [System.Text.StringBuilder]$sql=""
            foreach($sp in $CSVfile2.StoredProcedures){
             if($sp)
              {  $oRs_SP = $cmp2.GetBusinessObject(300)
                $sqlpath=$savepath2+$sp+'.sql'
                $sql =([System.Io.File]::OpenText($sqlpath)).ReadToEnd() 
                $oRs_SP.DoQuery($Sql) 
                Release-Ref ($oRs_SP)
              }
            
            }
            Write-Host  -ForegroundColor Green "End importing   stored procedures SQL......"
            Write-Host
            Write-Host 

#Enable sapb1 TN SP control
 Write-Host -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
$ticktNum =''
 fn_SAPB1_SP_control $ticktNum 'Y'

