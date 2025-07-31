cls
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1


$SourceSite    = "cstst"
# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp $SourceSite


 $csvfile = Import-Csv 'C:\Temp\KTFA.csv'

 foreach($r in $csvfile)
 {
     $oItm=$cmp.getbusinessobject(4)
 
     $oItm.ItemType = 3      # Fixed asset
     $oItm.SalesItem = 0
     $oItm.PurchaseItem = 0 

     $oItm.ItemCode =$r.'FA No.'
     $oItm.ItemName =$r.'FA Name'
     $oitm.AssetClass = $r.AssetClass
     $oItm.Location = 1 
     $oItm.AssetSerialNumber = 'George Xu'
     $oitm.AttributeGroups.Attribute1 = '2. 生产设备'

     #$oitm.DepreciationParameters.FiscalYear =  (Get-Date).Year
     #$oitm.DepreciationParameters.Add()
    # $oitm.DepreciationParameters.DepreciationArea = 100
     #$oitm.DepreciationParameters.DepreciationStartDate = '2021-11-30'
     # $oitm.DepreciationParameters.UsefulLife = $r.'Use life-Account(Month)'
     # $oitm.DepreciationParameters.DepreciationType = 'LINP'
     # $oitm.DepreciationParameters.Add()
     # $oitm.DepreciationParameters.DepreciationStartDate = '2021-11-30'
     # $oitm.DepreciationParameters.DepreciationArea = 200 
    #  $oitm.DepreciationParameters.UsefulLife = $r.'Use life-Tax(Month)'
    #  $oitm.DepreciationParameters.DepreciationType = 'LINP'

    $oItm.DistributionRules.DistributionRule ='PROD'
    $oItm.DistributionRules.DistributionRule2 = 'LOCAL'
    $oItm.DistributionRules.DistributionRule3 = 'EE'
    $oItm.DistributionRules.DistributionRule4 = 'COGP'
    $oItm.DistributionRules.DistributionRule5 = 'SIO2'
    $oItm.DistributionRules.ValidFrom  = '2021-11-30'
    $oItm.CapitalizationDate = '2021-11-30'

    Write-Host $r.'FA No.' $oItm.Add()  $cmp.GetLastErrorDescription()

    Release-Ref $oItm
 
     }


     # UPDATE Item group for Asset master data

      $csvfile = Import-Csv 'C:\Temp\KTFA.csv'

 foreach($r in $csvfile)
 {
     $oItm=$cmp.getbusinessobject(4)  
     if ($oItm.GetByKey($r.'FA No.'))
     {
       $oItm.ItemsGroupCode = $r.ITEMGROUP
        
       Write-Host $r.'FA No.' $oItm.update()  $cmp.GetLastErrorDescription()

       Release-Ref $oItm
     }
  }


  # Add Retirment doc

foreach($fa in $FAs){
                                  
      $oAssetService = $cmp.GetCompanyService().GetBusinessService(1470000094)   # AssetRetirementService 1470000094 
  $oAssetDocument = $oAssetService.GetDataInterface("0")                      
    $oAssetDocument.DocumentType = 6                                       # adtScrapping
     $oAssetDocument.SummerizeByProjects = 0     
     $oAssetDocument.SummerizeByDistributionRules = 1 
   
   $oline = $oAssetDocument.AssetDocumentLineCollection.Add()
   $oline.AssetNumber = $fa
   
 [void]$oAssetService.Add($oAssetDocument)

   Write-Host $fa  $cmp.GetLastErrorDescription()
   Release-Ref $oAssetService
   
}
   

# Update FA attributes

$oItm=$cmp.getbusinessobject(4)

$csv = Import-Csv \\sz-nas02\sznasit\dell\SZFA.csv

foreach($r in $csv){ 

IF ($oItm.GetByKey($r.ItemCode.TrimEnd())) {
 $oitm.AttributeGroups.Attribute1 = $r.TaxClass

 Write-Host $r.ItemCode  $oItm.Update() $cmp.GetLastErrorDescription()

 }
 }
