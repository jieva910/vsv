
# 2022.01.26
# update item price list 1 or 2,or other


$pricecsv = Import-Csv 'C:\temp\srprice5.csv' -Delimiter ","

foreach($p in $pricecsv){

if ($oitem.GetByKey($p.itemcod)) {

  $oItemprice = $oitem.PriceList 

 for($j=0;$j -lt $oItemprice.Count ;$j++)
 {
   $oItemprice.SetCurrentLine($j)
    if ($oItemprice.PriceList -eq $p.pircelist ) {
      # $oItemprice.PriceListName
       $oitemprice.Price = $p.price
       $oitemprice.Currency = $p.currency
       
       Write-Host $p.itemcod $oitem.update() $cmp.GetLastErrorDescription()
   } 
 
 }
}
}