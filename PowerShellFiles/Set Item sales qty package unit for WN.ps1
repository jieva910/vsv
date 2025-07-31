

. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$site = "wn"
$ticknum  = 'INC0193464'
  Fn_ConnectSAPB1 $site

     # Disable SP control 
  fn_SAPB1_SP_control $ticknum 'N' $site 

   $CSV = Import-Csv C:\Temp\WNITEM.csv



 $oItm=$cmp.getbusinessobject(4)

 FOREACH($R IN $CSV){

    if ( $oItm.GetByKey($R.'Item code'.Trim()) ) {
        $oItm.SalesQtyPerPackUnit = 1
         $outputcode = $oItm.Update()
    Write-Host $cmp.CompanyDB + ' '+ $R.'Item code' + ' updated  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
  
      }
    }

      
     # enable SP control
     $ticknum='' 
  fn_SAPB1_SP_control $ticknum 'Y'  $site 
