  
cls

$SourceSite    = "BY"

$ticktNum = 'INC0195792 '

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$CMP = New-Object -ComObject "SAPBOBSCOM.COMPANY"
# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $CMP $SourceSite

fn_SAPB1_SP_control $ticktNum 'N' $SourceSite

$oBP = $cmp.GetBusinessObject(2)



$CSV = IMPORT-CSV 'C:\Temp\SZBP.csv'
FOREACH($R IN $CSV)
{
 

 if ($oBP.GetByKey($R.'BP Code')){
  $oBP.Addresses.AddressType = 1
  $obp.addresses.AddressName = $R.Billto
  $obp.addresses.Country =$R.CountryS
  $obp.addresses.Street = $R.'Street/PO Box'
  $oBP.addresses.County = $obp.CardName
  $obp.addresses.Add()
   if(![string]::IsNullOrEmpty($r.IntercoSite)){$oBP.Password =$R.IntercoSite}
   if(![string]::IsNullOrEmpty($r.'Foreign Name')){$obp.CardForeignName=$r.'Foreign Name'}
  Write-Host "$($R.'BP Code') has been updated $($obp.Update())  $($cmp.GetLastErrorDescription())"
 }
 else {"$($r.'BP Code') not exist"}


}


