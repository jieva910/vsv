cls
# 2021.05.14 
# 初始化值
$site = 'sztst'

$ticknum = 'update BP housebank per krystal feng ticket'
$cmp = New-Object -ComObject "sapbobscom.company"
 # Load DI and connect to Company db
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1
Fn_ConnectSAPB1 $cmp $site


$cmp.XmlExportType  = 3 

$udo = 'Consumption Patterns'

$oCompServic = $cmp.GetCompanyService()
$oGeneralServic = $oCompServic.GetGeneralService($udo)
 $oGeneralData = $oGeneralServic.GetDataInterface(1)
$oGeneralParams  = $oGeneralServic.GetDataInterface([SAPbobsCOM.GeneralServiceDataInterfaces]::gsGeneralDataParams)  
  $oGeneralParams.setProperty("Code","202308")
  $oGeneralData = $oGeneralServic.GetByParams($oGeneralParams)

   $oGeneralData.ToXMLFile('c:\temp\x.xml')
