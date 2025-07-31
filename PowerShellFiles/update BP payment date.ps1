  
cls

$SourceSite    = "BY"

$ticktNum = 'INC0195792 '

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite

fn_SAPB1_SP_control $ticktNum 'N' $SourceSite

$oBP = $cmp.GetBusinessObject(2)
$oRS  = $cmp.GetBusinessObject(300)
$oRS.DoQuery("SELECT T0.[GroupNum] FROM OCTG T0 WHERE T0.[PymntGroup] ='Interco Payment'")

if(!$oRS.EoF){$paytermcode=$ors.Fields.item(0).value}



$s = gc C:\Temp\BYBP2.csv

for($i = 0; $i -lt $s.Count;$i++)
 {
   $r = $s[$i] -split ","
  
   
if ($oBP.GetByKey( $r[2])){

$obp.BPPaymentDates.PaymentDate = 23
#$oBP.PayTermsGrpCode = 3
Write-Host "$($r[2]) updated with err:” $oBP.Update() $cmp.GetLastErrorDescription()
}

}

$ticktNum =''
fn_SAPB1_SP_control $ticktNum 'Y' $SourceSite


