
$site = "SZ"
 $cmp = New-Object -ComObject "SAPBOBSCOM.Company"

    . C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

    fn_connectsapb1 $cmp $site

    
      $oitem = $cmp.getbusinessobject(4)

    $csv = Import-Csv 'C:\Dell\SGM ITEM.csv'

foreach($r in $csv)
 { 
    $oitem.ItemCode=$r.ItemCode
    $oitem.ItemName=$r.ItemName
    $oitem.ForeignName = $r.ForeignName
    $oitem.ItemsGroupCode = 143
    $oitem.PurchaseItem = 1
    $oitem.SalesItem = 1
    $oitem.InventoryItem =1 
    $oitem.ShipType  = $r.ShipType
    $oitem.ManageBatchNumbers =1
    $oitem.GTSItemSpec = $r.GTSItemSpec
    $oitem.GTSItemTaxCategory =$r.GTSItemTaxCategory
    $oitem.ProdStdCost = $r.'Production Std Cost '
    $oitem.Properties(3) = 1 
    $oitem.User_Text =$r.User_Text
    $oitem.DefaultWarehouse='SZ-B-QC'
    $oitem.Valid = 1
    $oitem.Frozen = 0
    $oitem.ValidFrom = '2025-07-01'
    $oitem.ValidTo = '2099-12-31'
    $oitem.PurchaseUnit = 'PCS'
    $oitem.SalesUnit = 'PCS'
    $oitem.InventoryUOM = 'PCS'
  #  $oitem.CostAccountingMethod = 'S'
    $oitem.UserFields.Fields.Item('U_PLN').value = $r.U_PLN 
     $oitem.UserFields.Fields.Item('U_PLS').value =  $r.U_PLS
      $oitem.UserFields.Fields.Item('U_PGR').value =  $r.U_PGR
       $oitem.UserFields.Fields.Item('U_Ves_RawMatPLN').value = $r.U_Ves_RawMatPLN 
        $oitem.UserFields.Fields.Item('U_Ves_FinPLN').value = $r.U_Ves_FinPLN
         $oitem.UserFields.Fields.Item('U_Ves_ItemType').value = $r.U_Ves_ItemType
          $oitem.UserFields.Fields.Item('U_prccode').value = $r.U_prccode
    		 $oitem.UserFields.Fields.Item('U_Ves_PRP1').value = $r.U_Ves_PRP1
     $oitem.UserFields.Fields.Item('U_Ves_PRP2').value = $r.U_Ves_PRP2
      $oitem.UserFields.Fields.Item('U_Ves_MfgLine').value = $r.U_Ves_MfgLine		
       write-host " $($r.ItemCode) added with err: $($oitem.Add()) $($cmp.GetLastErrorDescription()) " 
    				
    }

    # add warehouse to item code 

    $whsgrp = @('SZ-B-FG',
'SZ-B-QC',
'SZ-B-RM',
'SZ-B-SR',
'SZ-B-SG',
'SZ-B-SS',
'SZ',
'CT',
'WG',
'WS',
'WV',
'AS'
)

foreach($r in $csv)
 { 
    if ($oitem.GetByKey($r.ItemCode))
    {
     $oitem.DefaultWarehouse = 'SZ-B-QC'
     <# $i = 0 
      foreach($wh in $whsgrp)
      {   
          #  $oitem.WhsInfo.SetCurrentLine($i)      
             
              $oitem.WhsInfo.WarehouseCode =$wh
              $oitem.WhsInfo.add()   
            $i++  

         } #>
      Write-Host  "  $($r.ItemCode)  has been SET DEFAULT WAREHOUSE " $oitem.Update() $cmp.GetLastErrorDescription()
    
    }

}
