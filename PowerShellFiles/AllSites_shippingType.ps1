
<# Shipping Type Add,update,delete #>

cls

$sites = "sz","cs","kt","wn","as","by"

$csv = Import-Csv C:\temp\allsitesupplier.csv -Delimiter ","
$csv_shippingtype  = Import-Csv C:\Temp\newshippingtype.csv -Delimiter "," 

$ticktNum = "INC0275489"

FOREACH($site in $sites) # Loop site list
{
    $cmp = New-Object -ComObject "SAPBOBSCOM.Company"

    . C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

    fn_connectsapb1 $cmp $site

    $oShippingType = $cmp.GetBusinessObject(49)
    $obp = $cmp.GetBusinessObject(2)
    $ors = $cmp.GetBusinessObject(300)
     #disable SAPB1 TN sp control
     Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"

    fn_SAPB1_SP_control $ticktNum 'N' $site


     foreach($r in $csv_shippingtype ) # loop shipping type list  ,and add / update
    {
    $oShippingType.Website =$r.'Web Site'
    $oShippingType.Name =$r.TrnspName
    $oShippingType.UserFields.Fields.Item("U_VES_ShipCat").value ="B"

    Write-Host "$($site)..$($r.TrnspName) added with err code $($oShippingType.Add())..$($cmp.GetLastErrorDescription())"
    }
    
    # loop supplier list  ,and update shipping type
    foreach($r2 in $csv | Where-Object{$_.sitecode -eq $site} )  #  match to respective sitecode 
    {
      $billto = 0
      if ($oBP.GetByKey($r2.supplier))
      {   
        $ors.doquery("SELECT TrnspCode  FROM OSHP where TrnspName ='$($r2.shippingtype)'")
       if(!$ors.eof){ $obp.ShippingType =$ors.Fields.item(0).value}
   
         for ( $i =0 ;$i -lt $obp.Addresses.Count ;$i++ )
        {
           $oBP.Addresses.SetCurrentLine($i) 
         if ($oBP.Addresses.AddressType -eq 1) # only update billto country
         {   
                  
          $obp.addresses.Country=$r2.vendorcontry  
          $billto += 1
         }
         else 
         { $billto +=0 }
        }
  
       if (!$billto)                          # if no billto then add beginngin with line 0 
       {
         $oBP.Addresses.SetCurrentLine(0)
         $obp.addresses.AddressName=$r2.supplier
         $oBP.Addresses.AddressType = 1
        $obp.addresses.Country=$r2.vendorcontry
         $obp.addresses.Add()
       }
          

       Write-Host "$($site)..$($r2.supplier) updaed with err code $($obp.update())..$($cmp.GetLastErrorDescription())"
       }
    }

    #Enable sapb1 TN SP control
    Write-Host -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
    $ticktNum2 =''
    fn_SAPB1_SP_control $ticktNum2 'Y' $site
    Release-Ref $cmp
}