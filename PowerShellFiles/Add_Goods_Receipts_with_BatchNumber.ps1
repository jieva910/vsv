
<#
  Purpose : Add goods receipt with batch number 
  Date    : 2020.11
#>


cls

$cmp           = New-Object -ComObject 'SAPbobsCOM.Company'
$SourceSite    = "CSTST"



# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite


$oGoodsReceipt = $cmp.GetBusinessObject(59) # oInventoryGenEntry 59 

# Export Goods Receipts to XML file
$cmp.XmlExportType = 3 
if ($oGoodsReceipt.getbykey(122977)) 
{$oGoodsReceipt.SaveXML('c:\temp\goodsRecipt.xml')}

# add new goods receipt from XML
$New_GoodsReceipt = $cmp.GetBusinessObjectFromXML('c:\temp\goodsRecipt.xml',0)
$New_GoodsReceipt.add()
$cmp.GetLastErrorDescription()