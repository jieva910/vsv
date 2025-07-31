
cls
 # remove  code from Item master data
 # 2022.03.05.

$site = 'cs'
$ticknum = 'add new  and set GL advance determination rule and do stock revaluation'
$cmp = New-Object -ComObject "SAPBOBSCOM.COMPANY"

 # Load DI and connect to Company db
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1
Fn_ConnectSAPB1 $CMP $site

$oitm = $cmp.GetBusinessObject(4)

$oRs  = $cmp.GetBusinessObject(300)

$sql = "SELECT t.ItemCode FROM OITw t WHERE t.WhsCode = 'CS-GZ' AND t.ItemCode 	 <> '2026452852110CN@XX' AND t.AvgPrice = 0" #   has no existing trans"

$oRs.DoQuery($sql)

if (!$oRs.EoF)
{
  [xml]$Nodes = $oRs.GetAsXML()
	  
	  $Rows = $Nodes.SelectNodes("//row")
      $k = 1 
	  foreach($n in $Rows)
      {        
        if ($oitm.GetByKey($n.ItemCode))
        {
          $whsrows = $oitm.WhsInfo.Count
          for($i=0; $i -lt $whsrows;$i++)
          {
            $oitm.WhsInfo.SetCurrentLine($i)               # 物料主数据里面的仓库hang

            if ( $oitm.WhsInfo.WarehouseCode -eq  "CS-GZ")
            {
                $oitm.WhsInfo.Delete()                     # 删除指定的仓库code
                
                Write-Host -ForegroundColor Cyan "Process $($k)/$($Rows.Count) ..."

                Write-Host  " specific  of " $n.ItemCode " has been removed " $oitm.Update() $cmp.GetLastErrorDescription()
				
				break                                      # 退出当前循环
            }
  
          }
        }
          $k++
       }

}

Release-Ref $oitm