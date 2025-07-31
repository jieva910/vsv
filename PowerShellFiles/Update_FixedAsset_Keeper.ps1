
<#
  Purpose : Add New Document Numbering Series for New Posting Periods
  Date    : 2020.12
#>


cls

$SourceSite    = "BY"

$ticktNum = 'FA keeper update'



# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite


# Disable SP control
fn_SAPB1_SP_control $ticktNum 'N' $SourceSite

$starttime =Get-Date

$CSVfile = Import-Csv 'C:\Temp\AS FA.csv'

FUnction Update_FixedAsset_Keeper
{
  param ($FAitem,$FAkeeper)
  $oItm=$cmp.getbusinessobject(4)
    if ( $oItm.GetByKey($FAitem) ) {
        $oItm.AssetSerialNumber = $FAkeeper
         
       WRITE-HOST $cmp.CompanyDB  $FAitem  'updated  with error code:' $oItm.Update()  $cmp.GetLastErrorDescription()
    }
 }


 
 foreach($row in $CSVfile) {
     
  
          Update_FixedAsset_Keeper $row.AssetCode.Trim() $row.Keeper.Trim()
       
 }


    $endtime =Get-Date

    Write-Host -ForegroundColor Red ('running time is :' +($endtime-$starttime).totalseconds)
  
$ticktNum2 =''
fn_SAPB1_SP_control $ticktNum2 'Y' $SourceSite
