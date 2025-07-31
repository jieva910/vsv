
<#
  Purpose : Add goods receipt with batch number 
  Date    : 2020.11
#>


cls

$cmp           = New-Object -ComObject 'SAPbobsCOM.Company'
$SourceSite    = "sq"



# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite


$oCompServic = $cmp.GetCompanyService()
$oGeneralServic = $oCompServic.GetGeneralService('SWA_LD_TEXT')
  
# Add new code in existing UDO
 $csvfiel =  Import-Csv c:\temp\UDO_Code.csv
$oGeneralData = $oGeneralServic.GetDataInterface(1)  # GeneralData data interface

 foreach($row in $csvfiel)
 {      
  $oGeneralData.SetProperty("Code", $row.code.trim())
  $oGeneralData.SetProperty("U_Text",$row.U_Text)

  $oGeneralServic.add($ogeneraldata)

  [void]$cmp.GetLastError($errCode,$errMsg)

 }

  
# Add translation for UDO new code
 <#
    SELECT t.TranEntry,t.TableName,t.PK,t1.Trans FROM OMLT t INNER JOIN  MLT1 t1 ON t.TranEntry = t1.TranEntry
 
    WHERE t.pk LIKE '%USR%' AND t1.LangCode =15
 #>
$csvfile_langTranslation =  Import-Csv c:\temp\UDO_Code22.csv

$langTranslation = $cmp.GetBusinessObject(224)    #SAPbobsCOM.BoObjectTypes.oMultiLanguageTranslations
$langTranslation.FieldAlias = "U_Text"  #(userfieds)
$langTranslation.TableName = "@SWA_LD_TEXT" #(usertable)

foreach($row in $csvfile_langTranslation)
{ 
   $langTranslation.PrimaryKeyofobject = $row.code.Trim()
   $langTranslation.TranslationsInUserLanguages.LanguageCodeOfUserLanguage = 15
   $langTranslation.TranslationsInUserLanguages.Translationscontent =$row.translation

   Write-Host $langTranslation.add()  $cmp.GetLastErrorDescription()
   
 }



  # export specific row data to XML file 
   $cmp.XmlExportType = 3 
   $oGeneralParams = $oGeneralServic.GetDataInterface(3)  # gsGeneralDataParams = 3
   $oGeneralParams.SetProperty('Code','USRO2C039')
   $oGeneralData = $oGeneralServic.GetByParams($oGeneralParams)
   $ogeneraldata.ToXMLFile('c:\temp\udo_specific_row.xml')




 	



