# Specify the Execution times
$TriggerTimes = @(
    '5:30:00pm'
   )

# Sort in chronologic order
#  assuming the times format are the same
$TriggerTimes = $TriggerTimes | Sort-Object

function UpdateItemweight
{
 
    $ticknum = 'update item weight per audrey ticket'

    # load sapb1 di connection lib
    . C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

    $site = "SZtst"

    # SAPB1 DI connect to specific site
    Fn_ConnectSAPB1 $site
    fn_SAPB1_SP_control $ticknum 'N' $site 
    $starttime =Get-Date

    $csv = import-csv C:\Temp\PKITEMS.csv

    foreach($r in $csv)
    { 
        $oItm = $cmp.GetBusinessObject(4)
        if($oItm.GetByKey($r.CODE.Trim())){
          $oitm.SalesUnitWeight = $r.'UNIT_WEIGHT( kg ) '
          $oitm.PurchaseUnitWeight = $r.'UNIT_WEIGHT( kg ) '
    
        Write-Host $r.CODE 'update with err code:' $oItm.Update() $cmp.GetLastErrorDescription()
        }
    }

    $endtime =Get-Date

    Write-host 'running time:' ($endtime-$starttime).TotalSeconds 
    $ticknum=''
     fn_SAPB1_SP_control $ticknum 'Y' $site 
}

foreach ($t in $TriggerTimes)
{
    # Past time ?
    if((Get-Date) -lt (Get-Date -Date $t))
    {
        # Sleeping
        while ((Get-Date -Date $t) -gt (Get-Date))
        {
            # Sleep for the remaining time
            (Get-Date -Date $t) - (Get-Date) | Start-Sleep
        }

        # Trigger event
        #  insert your code here
        "# TriggerTime: '$t' - Executing my code here!"

        UpdateItemweight

    }else{"Belong to the past: '$t'"}
}