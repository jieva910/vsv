<#
  Purpose :  add new user in all sites
  Date    : 2021/7
#>

cls
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'

$ticktNum = "INC0254284"

function fn_updateVendorAddressName
{
   

$obp = $cmp.GetBusinessObject(2)
$ors = $cmp.GetBusinessObject(300)
$sql = "SELECT CardCode FROM OCRD 
WHERE CardCode in ('300350V',
'300503V',
'300531V',
'300770V',
'300928V',
'301030V',
'301036V',
'301188V',
'301330V',
'301702V',
'301714V',
'301715V',
'301723V',
'301724V',
'301725V',
'301745V',
'301746V',
'301761V',
'301762V',
'301763V',
'301774V',
'301777V',
'301860V',
'301927V',
'301933V',
'301942V',
'302213V',
'302384V',
'302976V',
'302984V',
'302987V',
'302995V',
'302997V',
'303017V',
'303052V',
'303073V',
'303116V',
'303132V',
'303183V',
'303207V',
'303247V',
'303265V',
'303272V',
'303350V',
'303469V',
'303491V',
'304759V',
'305086V',
'305265V',
'305307V',
'305328V',
'305397V',
'305578V',
'305623V',
'305678V',
'305841V',
'305843V',
'305965V',
'305978V',
'306075V',
'306162V',
'306175V',
'306187V',
'306292V',
'306321V',
'306497V',
'306504V',
'306517V',
'306546V',
'306666V',
'306678V',
'306754V',
'306763V',
'306804V',
'306852V',
'306853V',
'306861V',
'306862V',
'306957V',
'306958V',
'306993V',
'307024V',
'307084V',
'307169V',
'307260V',
'S00118',
'S00408',
'S01188',
'S01702',
'S01753',
'S01767',
'S02983',
'S02988',
'S03160')"

$ors.DoQuery($sql)

if (!$ors.EoF)
{
  $xmlDoc = [xml]$ors.GetAsXML()
  $xmlNodes = $xmlDoc.SelectNodes("//row")

  foreach($node in $xmlNodes)
  {
    if ($obp.GetByKey($node.CardCode))
    {
      $obp.Addresses.AddressName =$node.CardCode
      
      Write-Host "$($node.CardCode) has been updated addressname $($obp.update()) $($cmp.GetLastErrorDescription())"
    }
  
  
  }

}



}
 
function fn_AddEmailgrp ($site)  {
     $oComServc = $cmp.GetCompanyService()

$oEmailGrp=$oComServc.GetBusinessService(234000004) # EmailGroupsService 234000004 

$t =$oEmailGrp.GetDataInterface(0)

$t.EmailGroupCode = 'GRPO Not Invoiced'

$t.EmailGroupName = 'Send to supplier every month'
 
# write-host "New email group has been added with err code: " $oEmailGrp.Add($t) $cmp.GetLastErrorDescription()

$csv = Import-Csv C:\Temp\siteBP2.csv -Delimiter "," 
foreach($r in $csv |?{$_.site -match $site})
 {
    $obp = $cmp.GetBusinessObject(2)

    if($obp.GetByKey($r.cardcode))
    {  
       
    $i = $obp.ContactEmployees.Count
    $obp.ContactEmployees.Add()

   # contact person 对象有个bug (sapb1 9.2 pl08), 当没有联系人的时候，count = 1,当有1个联系人的时候，count 也 =1 ，
   # 为了区分，在这里检查第一个联系人名称，如果为空，则说明没有联系人，此时设置 setcurrentline = 0
    $obp.ContactEmployees.SetCurrentLine(0)
     if (!$obp.ContactEmployees.Name){  $obp.ContactEmployees.SetCurrentLine(0)    }
     else {$obp.ContactEmployees.SetCurrentLine($i)}
       $obp.ContactEmployees.Name = 'GRPO Not Invoiced'
       $obp.ContactEmployees.EmailGroupCode = 'GRPO Not Invoiced'
     write-host "$($site)-New contact employee of $($r.cardcode) has been added with err code: "  $obp.update() $cmp.GetLastErrorDescription()
       }
 }
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
 #AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
#BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="CORP\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
 CS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_CS";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
#KT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_KT";sapuser="CORP\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
#SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
#WE =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WE";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
#SQ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_SQ";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}  
#WN =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WN";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
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
         fn_AddEmailgrp $_
         #fn_updateVendorAddressName

         #Enable sapb1 TN SP control
         Write-Host -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =''
         fn_SAPB1_SP_control $ticktNum2 'Y' $_
        Write-Host "End processing ."       
       
        $cmp.Disconnect()

  } 

#Release-Ref($cmp)  