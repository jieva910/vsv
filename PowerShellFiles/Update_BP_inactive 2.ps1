# 2023.11
# Mass inactive supplier in 6 company db.

$sites = "by"

$CSVfile = Import-Csv C:\temp\allsitesupplier2.csv -Delimiter ","
$log = 'c:\temp\inativebp2.log'
$ticktNum = "INC0608123"

function fn_UpdateBP {
    param ($bpcode)
     $oBP=$cmp.getbusinessobject(2)
    if ( $oBP.GetByKey($bpcode) ) {
        $obp.valid = 0 
        $oBP.Frozen = 1
        #$oBP.FrozenFrom = get-date -Format "yyyy-MM-dd"
        #$oBP.FrozenTo = "2999-12-31"
        $oBP.FrozenRemarks = $ticktNum
       
        $outputcode = $oBP.Update()
        $outlog =$cmp.CompanyDB + ' '+ $bpcode + ' updated  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
        
      }
       else {$outlog = $bpcode + ' not exists'} 
       return $outlog
}

FOREACH($site in $sites  ) # Loop site list
{
    $cmp = New-Object -ComObject "SAPBOBSCOM.Company"

    . C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

    fn_connectsapb1 $cmp $site
     fn_SAPB1_SP_control $ticktNum 'N'  $site 

      ForEach ($row in $CSVfile | Where-Object{$_.sitecode -eq $site} ){           
                    try { $outlog2=fn_UpdateBP  $row.BPCode }
                    catch { $outlog2= $_.Exception.Message;continue }
                    $outlog2| Out-File -Append $log
                 }


    $ticknum2='' 
  fn_SAPB1_SP_control $ticknum2 'Y'  $site 

  $CMP.Disconnect()
}