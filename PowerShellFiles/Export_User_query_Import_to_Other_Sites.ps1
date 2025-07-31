<#
  Purpose : Export specific user query in sap and import to other sites ,make sure user query keep consistency among all SAP sites
  Date    : 2020.11
#>

cls
 
$cmp = New-Object -ComObject 'SAPbobsCOM.Company'
$cmp2 = New-Object -ComObject 'SAPbobsCOM.Company'
$Userqueryname = "Control Account by BP Detail(All Information)"
$ticktNum      = "INC0239955"
$xmlfile      = "c:\Temp\$Userqueryname.xml"
$SourceSite    = "SZTST"
$queryCateogry = "FMS Supplier Consng"

$xmlfile_copy = "c:\temp\$($Userqueryname)_copy.xml"

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp2 $SourceSite

 # Save user query as XML file

  $cmp2.XmlExportType = 3 
  $oUQ = $cmp2.GetBusinessObject(160) #Const oUserQueries = 160 (&HA0)
  $rs = $cmp2.GetBusinessObject('300') #recordset
  
 $querysql="SELECT t.IntrnalKey,t.QCategory,t1.CatName,t.QName FROM OUQR t INNER JOIN oqcn t1 ON t.QCategory = t1.CategoryId WHERE 
 t1.CatName like '%$queryCateogry%'"
    
 $rs.DoQuery($querysql)
 if (!$rs.EoF)
{
  $XML = [XML]$RS.GetAsXML()
  $NODES = $XML.SelectNodes("//row")

  foreach($n in $NODES)
  {
    $queryID =$n.IntrnalKey
    $sourceCategroyID =$n.QCategory
      If ($oUQ.GetByKey($queryID,$sourceCategroyID)){$oUQ.SaveXML("c:\temp\$($n.QName).xml") ;
    
  }
    
     write-host;Write-Host -ForegroundColor Green "Export User query finish"
     $cmp2.Disconnect()
      Write-Host
    }
}  



# ###############################################  Import User query #########################################
Function fn_ImportUserquery ($site)
{   
    $Target_queryCateogry = "[$site] $queryCateogry"
    $querysql2="SELECT t.CategoryId FROM OQCN t WHERE t.CatName ='$Target_queryCateogry'"
    
    [xml]$xmlUser_query=Get-Content -Path $xmlfile 
         
     $rs.DoQuery($querysql2)
          
      if (!$rs.EoF)
         {  $target_CategoryID=$rs.Fields.Item(0).value
            $xmlUser_query.BOM.bo.UserQueries.row.QueryCategory=$target_CategoryID.tostring()
            $xmlUser_query.Save($xmlfile_copy) 
         }

    # if user query exist then update it 
    $rs.DoQuery($querysql)
    if (!$rs.EoF)
    {
      $queryID=$rs.Fields.Item(0).value
      $CategroyID=$rs.Fields.Item(1).value
          If ($oUQ.GetByKey($queryID,$CategroyID))
          {     
              
               $oUQ.Browser.ReadXml($xmlfile_copy,0)
               Write-Host  "   Update User query with error code : "$oUQ.update() " and error description is :" $cmp.GetLastErrorDescription()
           }

     }
    
     #   user query not exist in target site then Add it
    else {

           $oUQ.Browser.ReadXml($xmlfile_copy,0)
           Write-Host  "Update User query with error code : "$oUQ.add() " and error description is :" $cmp.GetLastErrorDescription()
        } 
     
           
}

#  batch update user query within sites
$SAP_SiteConnS = @{  AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
CS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_CS";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
#KT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_KT";sapuser="CORP\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
#WG =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WG";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
WN =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WN";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
#WV =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WV";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
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
        else { Write-Host -ForegroundColor Green "01) " $cmp.CompanyDB connected successfully}

         #disable SAPB1 TN sp control
        Write-Host  -ForegroundColor Green "02) disable SAPB1 transaction notfication SP control"

            fn_SAPB1_SP_control $ticktNum 'N' $_

         Write-Host -ForegroundColor Cyan "03) Start update User query  ......"

            $cmp.XmlExportType = 3 
            $oUQ = $cmp.GetBusinessObject(160) #Const oUserQueries = 160 (&HA0)
            $rs = $cmp.GetBusinessObject('300') #recordset

           # Excute update uer query in each site
           fn_ImportUserquery $_


          Write-Host
         #Enable sapb1 TN SP control
         Write-Host -ForegroundColor Green "04) Enable SAPB1 transaction notfication SP control"
           $ticktNum2 = ""
           fn_SAPB1_SP_control $ticktNum2 'Y' $_

         Write-Host -ForegroundColor Cyan "05) End processing ."       
         Write-Host
         Write-Host

        $cmp.Disconnect()

  } 