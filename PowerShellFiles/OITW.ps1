



cls

<#
   设置仓库的最小库存。
#>

$SourceSite    = "SZTST"

$ticktNum = 'add goods reason code'

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite


$OITM  =$CMP.GetBusinessObject(4)
$oRs  = $CMP.GetBusinessObject(300)

$csv = Import-Csv C:\Temp\srmin.csv -Delimiter ","

foreach($itemcode in $csv)
{   
    if ($OITM.GetByKey($itemcode.'Item No.'.trim())) {
      # only upate new value
      if ($oitm.MinInventory -ne $itemcode.'Minimum stock' ) {
           if ($oitm.ManageStockBy = 1 )  # managed stock by 
           {
            $item = $itemcode.'Item No.'
            $sql = "with whse_line as (select ROW_NUMBER()over(order by whscode)-1 lineid,WhsCode from oitw t where t.itemcode = '$item' ) `
                   select lineid from whse_line where WhsCode = 'SR-W'"
             $oRs.doquery($sql)
             if(!$oRs.EoF){$line = $oRs.Fields.Item(0).value
                      $OITW = $OITM.WhsInfo
                      $OITW.SetCurrentLine($line)
                      $oitw.MinimalStock = $itemcode.'Minimum stock'
                      }
    
           }
           else {$oitm.MinInventory = $itemcode.'Minimum stock' }        # non managed by 
        write-host $itemcode.'Item No.' " update min stock " $itemcode.'Minimum stock' " with err code "    $OITM.Update();$CMP.GetLastErrorDescription()
      }
    }
}



