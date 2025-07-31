
<#
  Purpose : create PO
  Date    : 2020.12
#>


cls
$cmp = New-Object -ComObject 'sapbobscom.company'
$ticknum = ' '

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1


$site = "sz"

 Fn_ConnectSAPB1 $cmp $site




 # fn_SAPB1_SP_control $ticknum 'Y' $site


  ######### create PO

  # 1) Add single row PO 
$csv = Import-Csv -LiteralPath C:\Temp\SZPO.csv -Delimiter ','

foreach($row in $csv)
{
 $commts = $row.DocNum 
  $oPO  = $cmp.GetBusinessObject(22)
  $opo.CardCode = $row.CardCode
  $opo.DocDate = '2021-01-06'
  $opo.DocDueDate = '2021-12-31'
  $opo.UserFields.Fields.Item(’U_ShipTo‘).value = $row.'Customer Code'
  $opo.HandWritten = 0 
  $opo.Series  = 214   # pos21 in sz sapb1
  $opo.Confirmed = 1 
  $opo.Comments = "instead of  $commts for WG butterfly"
 
  $opo_line = $oPO.Lines

  $opo_line.ItemCode = $row.ItemCode
  $opo_line.Currency = 'RMB'
  $opo_line.UnitPrice  = $row.Price
  $oPO_line.Code = $row.WhsCode
  $OPO_LINE.Quantity =$row.Qty
  $OPO_LINE.UserFields.Fields.Item('U_Ves_Requisitioner').VALUE = 'Alice Ban'
   $OPO_LINE.UserFields.Fields.Item('U_Ves_PRApprover').VALUE = 'Carl Zhang'
   $OPO_LINE.CostingCode = $row.OcrCode			
   $OPO_LINE.CostingCode2 =$row.OcrCode2
    $OPO_LINE.CostingCode4 = $row.OcrCode4
     $OPO_LINE.CostingCode5 = $row.OcrCode5
 WRITE-HOST $commts $OPO.Add() $CMP.GetLastErrorDescription()
  $newPOentry = $cmp.GetNewObjectKey()

 if ($opo.GetByKey( $newPOentry ))
 {
  $opo.Printed = 1 
  Write-Host $newPOentry 'printed Status set yes'  $opo.Update() $cmp.GetLastErrorDescription()
 }
  Release-Ref $oPO 
}


  # 2) Add PO with multiple lines
 
$SZPO2 = Import-Csv 'C:\TEMP\BIGSCTPO.csv'
$GrpPO = $SZPO2 |Group-Object -Property 'Remark'
foreach($po in $GrpPO)
{  
  $oPO  = $cmp.GetBusinessObject(22)
  $opo.CardCode = '900499v'
  $opo.DocDate = get-date
  $opo.DocDueDate = '2022-12-31'
 # $opo.UserFields.Fields.Item(’U_ShipTo‘).value = $arrHead[2]
  $opo.HandWritten = 0 
  $opo.Series  = 280   # pos21 in sz sapb1
  $opo.Confirmed = 1 
  $opo.Comments = "INC0315917"
  $opo.DocumentsOwner = 206
  
  $opo_line = $oPO.Lines
  foreach($line in  $po.Group)
  {
    
   $opo_line.ItemCode = $line.ItemCode
   $opo_line.Currency = 'RMB'
   $opo_line.UnitPrice  = $line.Price
  # $oPO_line.Code = $line.WhsCode
   $OPO_LINE.Quantity =$line.Qty
   $OPO_LINE.UserFields.Fields.Item('U_Ves_Requisitioner').VALUE = 'Vincy Liu'
   $OPO_LINE.UserFields.Fields.Item('U_Ves_PRApprover').VALUE = 'Winnie Huang'
      $OPO_LINE.UserFields.Fields.Item('U_VES_BPShipDate').VALUE = '2023-12-30'
      $opo_line.ShipDate = '2023-06-30'
   $OPO_LINE.CostingCode = 'XX'			
   $OPO_LINE.CostingCode2 ='LOCAL'
    $OPO_LINE.CostingCode4 = 'OTCO'
    $OPO_LINE.CostingCode5 = 'OTHR'
     $opo_line.Add()
    
     $oPO.SpecialLines.LineType =[SAPbobsCOM.BoDocSpecialLineType]::dslt_Text
      $oPO.SpecialLines.AfterLineNumber = $OPO_LINE.LineNum-1     # note this value
      $oPO.SpecialLines.LineText   = $po.Name
      # $oPO.SpecialLines.SetCurrentLine($k)
       $oPO.SpecialLines.Add()
   
  }

  WRITE-HOST  $OPO.Add() $CMP.GetLastErrorDescription()
  if ($opo.GetByKey( $newPOentry ))
 {
  $opo.Printed = 1 
  Write-Host $opo.docnum ' printed Status set yes'  $opo.Update() $cmp.GetLastErrorDescription()
 }
  Release-Ref $oPO
}
 




