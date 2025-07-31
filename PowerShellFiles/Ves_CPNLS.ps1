
cls

$cmp = New-Object -ComObject "SAPBOBSCOM.Company"

$site = "sz"

. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

fn_connectsapb1 $cmp $site

$UserTable ='Ves_CPNLS'
 # $udt = $cmp.GetBusinessObject(153) #Const oUserTables = 153 (&H99)
 
$CSV = Import-Csv C:\Temp\CPNLS.csv -Delimiter ","

FOREACH($R IN $CSV)
{
  $UDT = $cmp.UserTables.Item($UserTable)

    IF($UDT.GetByKey($R.Code))
     {
  
       $UDT.UserFields.Fields("U_VES_TaxCOA2").value=$R.TaxCOA2
       $UDT.UserFields.Fields("U_VES_TaxCOA3").value =$R.TaxCOA3
      Write-Host "$($r.Code) has been update $($UDT.Update())..$($CMP.GetLastErrorDescription())"
     }

     Release-Ref $udt
 

}
  
