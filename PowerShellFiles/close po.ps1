
<#
  Purpose : close Open PO
  Date    : 2020.12
#>


cls

$ticknum = 'INC0319739'
$cmp = new-object -ComObject "sapbobscom.company"
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$csv = Import-Csv "C:\Temp\ChinaPO.csv" -Delimiter ","
$sites = "SZ","wn","BY","KT","AS"


function fn_ClosePo {
    param ($ponumber)
    $po = $cmp.GetBusinessObject(22)
    $rs = $cmp.GetBusinessObject('300') #recordset

    $rs.doquery(“select docentry from opor where docnum=" +$ponumber + "")
    if ($rs.EoF -eq $false)
     {  $poDocentry = $rs.Fields.Item(0).value
        if ($po.getbykey($poDocentry) -eq $true )
        {
          $outputcode = $po.CLOSE()
        $outlog =$cmp.CompanyDB + ' '+ $ponumber.tostring() + ' close  with error code:' +   $outputcode.tostring() + ' and error description is:' + $cmp.GetLastErrorDescription()
         return $outlog
         }
      }
}


foreach($site in $sites )
{
  # SAPB1 DI connect to specific site
  Fn_ConnectSAPB1 $cmp $site

  # Disable SP control 
  fn_SAPB1_SP_control $ticknum 'N'  $site 

  foreach($row in $csv | Where-Object{$_.site -eq $site} )
  {
    fn_ClosePo $row.DocNum
  }
 
  # Enable SP control
  $ticknum='' 
  fn_SAPB1_SP_control $ticknum 'Y'  $site 

  $CMP.Disconnect()
  Start-Sleep 5
}


