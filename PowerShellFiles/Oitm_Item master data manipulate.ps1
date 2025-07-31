
<#
  Purpose : Item master data manipulate
  Date    : 2021.1
#>


cls
$cmp = New-Object -ComObject 'sapbobscom.company'
$SourceSite    = "BY"
$ticknum   = 'Update item name'
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp $SourceSite 

# Disable SP control 
fn_SAPB1_SP_control $ticknum 'N'  $SourceSite 

 $oitemcode=$cmp.getbusinessobject(4)

 $csv_csitem =  import-csv C:\Temp\csitem.csv

 #  set default 

 foreach($r in $csv_csitem)
 {
    if ( $oitemcode.GetByKey($r.item) ) {
        $oitemcode.Default  = $r.dftwh
       
        $outputcode = $oitemcode.Update()
        $outlog =$cmp.CompanyDB + ' '+ $r.item + ' set default whse  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
        
      }
   else {$outlog = $r.item + ' not exists'} 

   write-output($outlog)
 }


 $csv_name=Import-Csv C:\Dell\BYitem.csv

 foreach($ritem in $csv_name)
 {
    if ( $oitemcode.GetByKey($ritem.PARENT_ITEM_CODE) ) {
        $oitemcode.ItemName = $ritem.'new CFOName'
         $oitemcode.ForeignName =$ritem.'new frgnname'
        $outputcode = $oitemcode.Update()
        $outlog =$cmp.CompanyDB + ' '+ $ritem.item + ' update item name  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
        
      }
   else {$outlog = $r.item + ' not exists'} 

   write-output($outlog)
 }