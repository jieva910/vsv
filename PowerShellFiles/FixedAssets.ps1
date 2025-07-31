

  
cls

$SourceSite    = "sztst"

$ticktNum = 'add goods reason code'

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite

$cmpSvc = $cmp.GetCompanyService()

<#
  DepreciationAreasService 1470000003   
FAAccountDeterminationsService 1470000002   
AssetClassesService 1470000032   
AssetCapitalizationService 1470000049   
AssetCapitalizationCreditMemoService 1470000060   
AssetTransferService 1470000090   
AssetManualDepreciationService 1470000075   
AssetRetirementService 1470000094 
AssetOriginalTypeEnum Enumeration  = aotRetirement
#>

$oAssetService= $cmpSvc.GetBusinessService(1470000094)  
$oAssetDocument  = $oAssetService.GetDataInterface(0)   #  adsAssetDocument
$oAssetDocument.SummerizeByDistributionRules = 1        #  Consolidate Journal Entry Rows by Distribution Rules
$oAssetLine = $oAssetDocument.AssetDocumentLineCollection
$oJournalEN = $oAssetDocument.AssetDocumentAreaJournalCollection.Add()

$oJournalEN.DepreciationArea = "*"
# $oAssetDocument.DocumentType = 6                        # AssetDocumentTypeEnum = adtSales
$oAssetLine1=$oAssetLine.add()
$oAssetLine1.AssetNumber = 'MR00433'
$oAssetLine2=$oAssetLine.Add()
$oAssetLine2.AssetNumber = 'VO0047'




$oAssetService.add($oAssetDocument)

$CMP.GetLastErrorDescription()


# Export FA retirement to xml 
$ofaDocumentParams  = $oAssetService.GetDataInterface(2)
$ofaDocumentParams.Code  = 36
$oAssetDoc = $oAssetService.Get($ofaDocumentParams)
$oAssetDoc.ToXMLFile('c:\temp\FA3.XML')