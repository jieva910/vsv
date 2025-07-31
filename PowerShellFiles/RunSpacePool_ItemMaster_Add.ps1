
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$site = 'SZTST'
Fn_ConnectSAPB1 $site


$starttime = Get-Date

$SessionState = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$Pool = [runspacefactory]::CreateRunspacePool(1, [int]$env:number_of_processors, $SessionState, $Host)
$Pool.Open()
 

$ScriptBlock = {
  
   Param(
      $oCmp,
  $ItemCode,
  $ItemName,
  $ForeignName,
  $ItemsGroupCode,
  $PurchaseItem,
  $InventoryItem,
  $SalesItem,
  $ManageBatchNumbers,
  $Properties,
  $InventoryUOM,
  $SalesUnitHeight,
  $SalesUnitLength,
  $SalesUnitWidth,
  $U_PLN,
  $U_PLS,
  $U_Ves_FinPLN
   )
    $oItem=$oCmp.getbusinessobject(4) 

    $oItem.ItemCode = $ItemCode
    $oItem.ItemName = $ItemName
    $oitem.ItemsGroupCode = $ItemsGroupCode
    $oitem.PurchaseItem =$PurchaseItem 
    $oitem.InventoryItem =  $InventoryItem
    $oitem.SalesItem =$SalesItem 
    $oItem.ForeignName = $ForeignName
    $oItem.ManageBatchNumbers = $ManageBatchNumbers 
    $oItem.Properties(1) = $Properties 
    $oItem.InventoryUOM = $InventoryUOM
    $oItem.SalesUnit = 10 
    $oItem.SalesUnitHeight = $SalesUnitHeight
    $oItem.SalesUnitLength = $SalesUnitLength
    $oItem.SalesUnitWidth = $SalesUnitWidth 

    $oItem.UserFields.Fields.Item("U_PLN").Value = $U_PLN
    $oItem.UserFields.Fields.Item("U_PLS").Value = $U_PLS
    $oItem.UserFields.Fields.Item("U_Ves_FinPLN").Value = $U_Ves_FinPLN

    $oItem.WhsInfo.Code='SZ-A-FG'
    $oItem.WhsInfo.Add()
    $oItem.WhsInfo.Code='SZ-A-RT'
    $oItem.WhsInfo.Add()
    WRITE-OUTPUT ($oItem.Add() + $oCmp.GetLastErrorDescription())
}

 
$threads = @()
$csv_item = Import-Csv 'C:\Temp\SZITEM.csv'
$handles =foreach ($r in $csv_item)
{
  $Paralist = @{
  oCmp = $cmp
  ItemCode=$r.'ItemCode '
  ItemName=$r.'ItemName '
  ForeignName=$r.'ForeignName '
  ItemsGroupCode=$r.'ItemsGroupCode '
  PurchaseItem=$r.'PurchaseItem '
  InventoryItem=$r.'InventoryItem '
  SalesItem=$r.'SalesItem '
  ManageBatchNumbers=$r.'ManageBatchNumbers '
  Properties=$r.'Properties(1) '
  InventoryUOM=$r.'InventoryUOM '
  SalesUnitHeight=$r.'SalesUnitHeight '
  SalesUnitLength=$r.'SalesUnitLength '
  SalesUnitWidth=$r.'SalesUnitWidth '
  U_PLN=$r.U_PLN
  U_PLS=$r.U_PLS
  U_Ves_FinPLN=$r.U_Ves_FinPLN
  }

  $powershell = [powershell]::Create()
  [void]$powershell.AddScript($ScriptBlock)
  [void]$powershell.AddParameters($Paralist)
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
  if (-not $done) { Start-Sleep -Milliseconds 500 }
} until ($done)
 


$endtime =Get-Date

Write-output ("Running Time is :" + ($endtime-$starttime).TotalSeconds)

