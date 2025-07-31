    
   $CMP = New-Object  -ComObject "SAPBOBSCOM.COMPANY"

  . C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

    FN_CONNECTSAPB1 $CMP KT 

    (Get-Culture).TextInfo.ListSeparator  # check OS culture for CSV delimiter
    $DecimalSeparator = (Get-Culture).NumberFormat.NumberDecimalSeparator # . for china  , for EMEA country

     $oitem = $cmp.getbusinessobject(4)

    $csv = Import-Csv C:\Temp\GHnetweights.csv

    foreach($c in $csv){$C}
    {
       
     if ($oitem.GetByKey($c.ItemCode)){
        if ( $DecimalSeparator -eq "." ) {
        $oitem.InventoryWeight = $c.'Inventory Weight'
        $oitem.PurchaseUnitWeight = $c.'Purchasing Weight'
        $oitem.SalesUnitWeight  = $c.'Sales Weight'
        }
      else {
       # for EMEA Country 
        $oitem.InventoryWeight = $c.'Inventory Weight'.Replace(".",",")
        $oitem.PurchaseUnitWeight = $c.'Purchasing Weight'.Replace(".",",")
        $oitem.SalesUnitWeight  = $c.'Sales Weight'.Replace(".",",")
        
      
      }

        write-host  $c.ItemCode  $oitem.Update()    $cmp.GetLastErrorDescription()
     
         }

    
    
    }


    $CSV_KT=Import-Csv C:\Dell\KTITEM.csv

    FOREACH($C IN $CSV_KT)
    {
       if ($oitem.GetByKey($c.ItemCode)){
        $oitem.ForeignName =$C.FrgnName
        Write-Host $C.ItemCode $oitem.UPDATE() $CMP.GetLastErrorDescription()
       
       }
    }

   