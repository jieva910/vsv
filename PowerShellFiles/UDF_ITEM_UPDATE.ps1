$CMP  = New-Object -COMOBJECT "SAPBOBSCOM.COMPANY"

. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

Fn_ConnectSAPB1 $CMP "SZ"

$OITEM = $CMP.GetBusinessObject(4)


$CSV = Import-Csv C:\Temp\PrimaryMix-mapping-SZ-20220210.csv


FOREACH($R IN $CSV)
{
  IF ($OITEM.GetByKey($R.ItemCode))
  {
   $OITEM.UserFields.Fields.Item("U_Ves_Mix").VALUE= $R.U_Ves_Mix

   Write-Host "UPDATE ITEM $($R.ItemCode) " $OITEM.UPDATE() $CMP.GetLastErrorDescription()
  }

}