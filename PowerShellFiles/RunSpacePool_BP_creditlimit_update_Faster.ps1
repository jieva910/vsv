
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$site = 'SZTST'
Fn_ConnectSAPB1 $site


$starttime = Get-Date

$throttleLimit = 4
$SessionState = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$Pool = [runspacefactory]::CreateRunspacePool(1, [int]$env:number_of_processors, $SessionState, $Host)
$Pool.Open()
 
$ScriptBlock = {
    param($cmp,$CardCode,$creditline,$DebtLine)
 
   # Start-Sleep -Milliseconds 200
   # "Done processing ID $id"
     $oBP=$cmp.getbusinessobject(2)
    if ( $oBP.GetByKey($CardCode) ) {
        $oBP.CreditLimit =$creditline
        $oBP.MaxCommitment = $DebtLine
        $outputcode = $oBP.Update()
        $outlog =$cmp.CompanyDB + ' '+ $CardCode + ' updated  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
         $outlog
      }

}
 
$threads = @()

$CSV = Import-Csv 'C:\Temp\SZBP.csv'
$handles = foreach ($x in $csv) {  
    $powershell = [powershell]::Create()
    [void]$powershell.AddScript($ScriptBlock).AddArgument($cmp)
    [void]$powershell.AddArgument($x.CardCode)
    [void]$powershell.AddArgument($x.CreditLine)
    [void]$powershell.AddArgument($x.DebtLine)
    $powershell.RunspacePool = $Pool
    $powershell.BeginInvoke()
  $threads += $powershell
}
 
do {
  $i = 0
  $done = $true
  foreach ($handle in $handles) {
    if ($handle -ne $null) {
      if ($handle.IsCompleted) {
        $threads[$i].EndInvoke($handle)
        $threads[$i].Dispose()
        $handles[$i] = $null
      } else {
        $done = $false
      }
    }
    $i++
  }
  if (-not $done) { Start-Sleep -Milliseconds 200 }
} until ($done)

$endtime =Get-Date

Write-output ("Running Time is :" + ($endtime-$starttime).TotalSeconds)

