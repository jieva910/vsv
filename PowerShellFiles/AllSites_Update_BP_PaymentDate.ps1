


cls

$ticknum = 'INC0443542'

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$csv = Import-Csv 'C:\Dell\Payment term change list.csv' -Delimiter ","
$sites ="sz" #,,"sz",,"kt"


function fn_UpdateBPPaymentDate {
    param ($Bpcode)
    $oBP = $cmp.GetBusinessObject(2)

        if ($oBP.getbykey($Bpcode))
        {
          $oBP.BPPaymentDates.PaymentDate = 23
          $outputcode = $oBP.update()
        $outlog =$cmp.CompanyDB + ' '+ $Bpcode+ ' update payment date  with error code:' +   $outputcode.tostring() + ' and error description is:' + $cmp.GetLastErrorDescription()
         return $outlog
         }
      }


function fn_UpdateBPPCity {
    param ($Bpcode)
    $oBP = $cmp.GetBusinessObject(2)

        if ($oBP.getbykey($Bpcode))
        {
          
          $outputcode = $oBP.update()
        $outlog =$cmp.CompanyDB + ' '+ $Bpcode+ ' update payment date  with error code:' +   $outputcode.tostring() + ' and error description is:' + $cmp.GetLastErrorDescription()
         return $outlog
         }
      }


function fn_UpdateSupplierPaymentTerm {
    param ($Bpcode,$groupnum)
    $oBP = $cmp.GetBusinessObject(2)
    $ors = $cmp.GetBusinessObject(300)
    $ors.doquery("SELECT t.GroupNum FROM octg t WHERE t.PymntGroup IN ('$($groupnum)')")
        if ($oBP.getbykey($Bpcode) -and !$ors.Eof)
        {
          $oBP.PayTermsGrpCode =$oRs.Fields.Item(0).value
          $outputcode = $oBP.update()
        $outlog =$cmp.CompanyDB + ' '+ $Bpcode+ ' update payment term  with error code:' +   $outputcode.tostring() + ' and error description is:' + $cmp.GetLastErrorDescription()
         return $outlog
         }
      }

foreach($site in $sites)
{
  # SAPB1 DI connect to specific site
  $cmp = New-Object -ComObject 'sapbobscom.company'
  Fn_ConnectSAPB1  $cmp $site

  # Disable SP control 
  fn_SAPB1_SP_control $ticknum 'N'  $site 

  foreach($row in $csv  | Where-Object{$_.site -eq $site} )
  {
    
    fn_UpdateSupplierPaymentTerm $row.CardCode.Trim() $row.'New Payment Term'.Trim()
  }
 
  # Enable SP control
  $ticknum2='' 
  fn_SAPB1_SP_control $ticknum2 'Y'  $site 

  $CMP.Disconnect()
  Start-Sleep 3
}


