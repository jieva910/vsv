
cls

$cmp = New-Object -ComObject 'SAPBOBSCOM.COMPANY'
$SourceSite    = "CS"

$ticknum       = 'INC0144985'
# load sapb1 di connection lib


.  C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp $SourceSite

 
# Disable SP control 
fn_SAPB1_SP_control $ticknum 'Y'  $SourceSite 



 $oSpecialPrices = $cmp.GetBusinessObject(7)    # oSpecialPrices = 7

 # remove special price for specific item and BP

 $CS_CSV = Import-Csv 'C:\Temp\CSBP.csv' -Delimiter ","

 FOREACH($R IN $CS_CSV)
 { 
  
    if ($oSpecialPrices.GetByKey($R.olditem,$R.BP))
         {
         Write-Host  $R.olditem $R.BP  $oSpecialPrices.Remove() $cmp.GetLastErrorDescription()
          
         }
  }


  # remove from  SPP1
   FOREACH($R IN $CS_CSV)
 { 
  
    if ($oSpecialPrices.GetByKey($R.olditem,$R.BP))
         {
        
           $oSpecialPrices.SpecialPricesDataAreas.SetCurrentLine(0)
              $oSpecialPrices.SpecialPricesDataAreas.Delete()
         Write-Host  $R.olditem $R.BP  $oSpecialPrices.update() $cmp.GetLastErrorDescription()
          
         }
  }





#   add new item in OSPP FOR cs
 FOREACH($R2 IN $CS_CSV)
 { 
    $oSpecialPrices.ItemCode = $R2.newitem
   $oSpecialPrices.CardCode = $R2.BP
     $oSpecialPrices.Price =$R2.specialprice
    
  Write-Host  $R2.newitem $R2.BP $oSpecialPrices.Add() $cmp.GetLastErrorDescription()
   
 
  }




 # update special price to date.

 $CSV = Import-Csv  'C:\Temp\OSPRICE.csv'

 FOREACH($ROW IN $CSV)
 {
  
 if ($oSpecialPrices.GetByKey($ROW.ItemCode,$ROW.CardCode))
 { 
   
     $oSpecialPrices.SpecialPricesDataAreas.SetCurrentLine($ROW.LINENUM)
     $oSpecialPrices.SpecialPricesDataAreas.Dateto = '2021-01-31'
 
   Write-Host $oSpecialPrices.Update() $cmp.GetLastErrorDescription()
 }

 }


 

 # update special price header

 $CSV = Import-Csv  'C:\Temp\CSBP.csv'

 FOREACH($ROW IN $CSV)
 {
  
 if ($oSpecialPrices.GetByKey($ROW.ItemCode,$ROW.CardCode))
 { 
   
     $oSpecialPrices.Price = $ROW.specialprice
     
   Write-Host $ROW.ItemCode $ROW.CardCode $oSpecialPrices.Update() $cmp.GetLastErrorDescription()
 }

 }