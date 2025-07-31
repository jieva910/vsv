
$cmp  = new-object -ComObject "sapbobscom.company"

$cmp.Server = "SZ-SAPSTG91"
$cmp.SLDServer ="SZ-TSTSAPLIC92:40000"
$cmp.CompanyDB = "SAPB1_sz_TST"
$cmp.DbServerType = 10
$cmp.DbUserName ="b1if"
$cmp.DbPassword="Vsvapp@202333"
$cmp.UserName="montova"
$cmp.Password="ButterSZ"

$cmp.Connect()|Out-Null
$cmp.GetLastErrorDescription()




    # 定义分组维度（使用有序哈希表保证顺序）
$groupDimensions = [ordered]@{
    Header = 'DraftCreate','CloseDelivery','Comments','Reference2',
            'DocDate','TaxDate','PaymentGroupCode','DocEntry'
    
    Line = 'itemcode','quantity','UnitPrice','Currency','AccountCode',
          'CostingCode','CostingCode2','CostingCode3','CostingCode4','CostingCode5'
    
    Batch = 'batchnumber','BatchQuantity','Batch_U_FinUse'
}

$logFile = "C:\Logs\SAP_Consignment_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
function Log-Message {
    param($Message)
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
}

 

Try {
 $oRs = $cmp.getbusinessobject(300)
 $oRs.doquery("SELECT *,format(BatchManufacturingDate,'yyyy-MM-dd') BatchManufacturingDate2 FROM VES_AI_ConsignmentService")
 if ( $oRs.EoF) { Log-Message 'no record found';return}
 $xml =[xml]$oRs.getasxml()
 $nodes = $xml.selectnodes("//row")

# 主分组逻辑
$nodesGrp = $nodes | Group-Object $groupDimensions.Header

foreach($headerGroup in $nodesGrp) {
    "Header Group: $($headerGroup.Name)"
     $sampleItem = $headerGroup.Group[0]
    $cmp.StartTransaction()
  try {
  $oOIGN = $cmp.getbusinessobject(59)
  $oOIGN.DocDate = get-date 
  $oOIGN.TaxDate = Get-Date -Format "yyyy-MM-dd"  #($n.Name -split ",")[4]
  $oOIGN.Reference2 =$sampleItem.Reference2
  $oOIGN.Comments = $sampleItem.Comments
  $oOIGN.PaymentGroupCode =$sampleItem.PaymentGroupCode
    $lineGroups = $headerGroup.Group | Group-Object $groupDimensions.Line
     $i=0
    foreach($lineGroup in $lineGroups) {
        "`tLine Group: $($lineGroup.Name)"
        $linitem = $lineGroup.Group[0] 
        $oOIGN.Lines.ItemCode = $linitem.ItemCode
        $oOIGN.Lines.Quantity = $linitem.Quantity   
          $oOIGN.Lines.UnitPrice = $linitem.UnitPrice
          $oOIGN.Lines.Currency = $linitem.Currency
          $oOIGN.Lines.WarehouseCode='CT'
          $oOIGN.Lines.AccountCode=$linitem.AccountCode
          $oOIGN.Lines.CostingCode=$linitem.CostingCode
          $oOIGN.Lines.CostingCode2=$linitem.CostingCode2
          $oOIGN.Lines.CostingCode3=$linitem.CostingCode3
          $oOIGN.Lines.CostingCode4=$linitem.CostingCode4
          $oOIGN.Lines.CostingCode5=$linitem.CostingCode5
          $oOIGN.Lines.UserFields.Fields.Item('U_manufsource').Value = $linitem.ManufSource
          $oOIGN.Lines.UserFields.Fields.Item('U_FinUse').Value=$linitem.U_FinUse
        $batchGroups = $lineGroup.Group | Group-Object $groupDimensions.Batch
          $j=0 
        $batchGroups | ForEach-Object {
            "`t`tBatch Group: $($_.Name)"
            $batch=$_.Group[0]
             if ($batch.BatchNumber.Length -gt 36)
          {
           $oOIGN.Lines.BatchNumbers.BatchNumber =[string]$batch.ReferencedDocNumber+[string]$batch.LineNum+"_"+$batch.U_FinUse
          }
          else {  $oOIGN.Lines.BatchNumbers.BatchNumber = $batch.BatchNumber }
          $oOIGN.Lines.BatchNumbers.AddmisionDate=$batch.BatchManufacturingDate2
          $oOIGN.Lines.BatchNumbers.ManufacturingDate =$batch.BatchManufacturingDate2
          $oOIGN.Lines.BatchNumbers.Quantity = $batch.BatchQuantity
          $oOIGN.Lines.BatchNumbers.UserFields.Fields.Item('U_FinUse').Value = $batch.U_FinUse
   
            $oOIGN.Lines.BatchNumbers.SetCurrentLine($j)
               $oOIGN.Lines.BatchNumbers.Add()
           $j++
        }
        
     $oOIGN.Lines.SetCurrentLine($i) 
      $oOIGN.lines.add()
       $i++
  $rtcode = $oOIGN.Add()
  if($rtcode){ Log-Message "$($cmp.GetLastErrorDescription())";return} 
  $oDln = $cmp.GetBusinessObject(15)
  if($oDln.GetByKey($sampleItem.DocEntry))
    {$rtcode = $oDln.Close()
    if($rtcode){ Log-Message "Failed to close ODLN: $($cmp.GetLastErrorDescription())" ; return} 
    }
  $cmp.EndTransaction(0)
      Log-Message "Successfully created OIGN and closed ODLN for DocEntry: $($sampleItem.DocEntry)"
   }


  }
 catch {
            $cmp.EndTransaction(1)
            Log-Message "Transaction rolled back for DocEntry: $(($n.Name -split ",")[7])). Error: $($_.Exception.Message)"
            return
        }
   }
 }
catch {
    Log-Message "Script failed: $($_.Exception.Message)"
    return
}
finally {
        # 14. 对象释放
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($cmp) | Out-Null
           }