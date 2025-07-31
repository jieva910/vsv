<#
  Purpose : close Open PO
  Date    : 2020.12
#>


cls
$cmp = New-Object -ComObject "sapbobscom.company"
$ticknum = 'INC0155397'

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$site = "cs"

  Fn_ConnectSAPB1 $cmp $site

  # Disable SP control 
  fn_SAPB1_SP_control $ticknum 'N'  $site 

   
$csv = Import-Csv "C:\Temp\csbpcredit.csv" -Delimiter ","
$oBP=$cmp.getbusinessobject(2)

function fn_getPaymnttrmcode ([string]$p)
{
$ors = $cmp.GetBusinessObject(300)

$sql = "SELECT t.GroupNum FROM OCTG t WHERE LEFT(t.PymntGroup,2) = '$($p)'"

$ors.DoQuery($sql)

 return $ors.Fields.Item(0).value 

}



foreach($row in $csv)
{
  
 if ( $oBP.GetByKey($row.CustomerCode) ) {

     if ($row.'Customer Stauts' -eq'Active') {
     
       $oBP.CreditLimit =[INT]$row.FinalLimit
        $oBP.MaxCommitment = [INT]$row.FinalLimit
         $oBP.PayTermsGrpCode =fn_getPaymnttrmcode  $row.paycode    
     
     }
     elseif ($row.'Customer Stauts' -eq 'Inctive')
     {   $oBP.Frozen = 1
        $oBP.Valid = 0 
       }   

         Write-Host $cmp.CompanyDB  $row.CustomerCode  ' updated  with error code:'    $oBP.Update()  ' and error description is:' + $cmp.GetLastErrorDescription() 
      }

}

    


    if ( $oBP.GetByKey('122206') ) {
        $oBP.CreditLimit =0
        $oBP.MaxCommitment = 0
         $oBP.PayTermsGrpCode =45 
        Write-Host $cmp.CompanyDB   ' updated  with error code:'    $oBP.Update()  ' and error description is:' + $cmp.GetLastErrorDescription() 
     }
