

cls

$ticknum = 'INC0319739'
$cmp = new-object -ComObject "sapbobscom.company"
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1
Fn_ConnectSAPB1 $cmp 'sz'
$csv = Import-Csv 'C:\Temp\new PO owner.csv' -Delimiter ","

$opo=$cmp.GetBusinessObject(22)


foreach($c in $csv)
{
 $opo.GetByKey($c.docentry)
 $opo.DocumentsOwner = 257

 Write-Host "$($c.docentry) $($opo.Update()) upated $($cmp.GetLastErrorDescription()) "

}

