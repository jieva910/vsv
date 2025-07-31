

. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$site = "cs"
$ticknum  = 'INC0193464'
  Fn_ConnectSAPB1 $site

     # Disable SP control 
  fn_SAPB1_SP_control $ticknum 'N' $site 

   $CSV = Import-Csv C:\Temp\copycs.csv



 $oItm=$cmp.getbusinessobject(4)

 FOREACH($R IN $CSV){
   
   try {
    if ( $oItm.GetByKey($R.PatternNumber.Trim()) ) {
        IF ( $oItm.ForeignName -ne   $oitm.GTSItemSpec)
        {
           $oitm.GTSItemSpec = $oItm.ForeignName
                     
        }
      }
     }
    catch {$msg=$_.exception}

    finally{  Write-Host $cmp.CompanyDB  $R.PatternNumber ' updated  with exception:'  $($msg.Message)  " with error code:" $oItm.Update()  ' and error description is:'  $cmp.GetLastErrorDescription()  }
 }

      
     # enable SP control
     $ticknum='' 
  fn_SAPB1_SP_control $ticknum 'Y'  $site 
