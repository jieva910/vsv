cls


$cmp_source = New-Object -COMObject 'SAPbobsCOM.Company'
$cmp_target = New-Object -COMObject 'SAPbobsCOM.Company'

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$SourceSite    = "sq"
$targetSite = "cstst"
$ticktNum = "INC0223469"

# SAPB1 DI connect to source site
Fn_ConnectSAPB1 $cmp_source $SourceSite


# SAPB1 DI connect to target site
Fn_ConnectSAPB1 $cmp_target $targetSite

# Disable SP control
fn_SAPB1_SP_control $ticktNum 'N' $targetSite


# source site company service
$comserv = $cmp.GetCompanyService()

# FAAccountDeterminationsService 1470000002 
$oFAAccountDeterminationsService = $comserv.GetBusinessService(1470000002)
   
# Target site company service
$comserv_target = $cmp_target.GetCompanyService()

# FAAccountDeterminationsService 1470000002 
$oFAAccountDeterminationsService_target = $comserv_target.GetBusinessService(1470000002) 
  
$FAacctdetermcodes = 1041,1040,1051,1044

foreach($code in $FAacctdetermcodes)
{   
    # Retrieve FA account determination rules from Source site
    $s1 = $oFAAccountDeterminationsService.GetDataInterface(2)  # FAAccountDeterminationParams 
    $s1.Code = $code
    $s = $oFAAccountDeterminationsService.Get($s1)
    $strXml = $s.ToXMLString()

     # import FA account determination rules into Target site
     $t1 = $oFAAccountDeterminationsService_target.GetDataInterfaceFromXMLString($strXml)
      $oFAAccountDeterminationsService_target.Add($t1)
    $cmp_target.GetLastErrorDescription()
   
}




# source site company service for Asset class
$oCompanyService = $cmp.GetCompanyService()
# AssetClassesService  
$oAssetClassesService = $oCompanyService.GetBusinessService(1470000032)   

# Target site company service for Asset class
$oCompanyService_target = $cmp_target.GetCompanyService()
 # AssetClassesService  
$oAssetClassesService_target = $oCompanyService_target.GetBusinessService(1470000032)  

$assetClasses = "NC-CAP-BT 3rd PRT","NC-CAP-FO 3rd PRT","NC-CAP-MO 3rd PRT"

foreach($class in $assetClasses)
{   
    # Retrieve asset class from Source site
    $sClass1 = $oAssetClassesService.GetDataInterface(2)  # assectclassParams 
    $sClass1.Code = $class
    $sClas = $oAssetClassesService.Get($sClass1)
    $strXmlassetClass = $sClas.ToXMLString()

     # import asset class into Target site
     $tclass1 = $oAssetClassesService_target.GetDataInterfaceFromXMLString($strXmlassetClass)
      $oAssetClassesService_target.Add($tclass1)
    $cmp_target.GetLastErrorDescription()
   
}


