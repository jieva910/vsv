
<#
  Purpose : Item master data manipulate USING xml EXPORT AND UPDATE WHEN ADDING OR DELETING XML ELEMENT.
  Date    : 2021.3
#>


cls

$SourceSite    = "sztst"
$ticknum   = 'INC0146528'
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite

# Disable SP control 
fn_SAPB1_SP_control $ticknum 'N'  $SourceSite 


function addElement($e1, $name2, $value2, $attr2)
{
    if ($e1.gettype().name -eq "XmlDocument") {$e2 = $e1.CreateElement($name2)}
    else {$e2 = $e1.ownerDocument.CreateElement($name2)}
    if ($attr2) {$e2.setAttribute($value2,$attr2)}
    elseif ($value2) {$e2.InnerText = "$value2"}
    return $e1.AppendChild($e2)
}

function formatXML([xml]$xml)
{
    $sb = New-Object System.Text.StringBuilder
    $sw = New-Object System.IO.StringWriter($sb)
    $wr = New-Object System.Xml.XmlTextWriter($sw)
    $wr.Formatting = [System.Xml.Formatting]::Indented
    $xml.Save($wr)
    return $sb.ToString()
}

$xml = New-Object system.Xml.XmlDocument
$xml1 = addElement $xml "ItemPreferredVendors"
$xml2 = addElement $xml1 "row"
$xml3 = addElement $xml2 "BPCode" "900010v"
#$xml3 = addElement $xml2 "d" "attrib" "attrib_value"

write-host `nFormatted XML:`r`n`n(formatXML $xml.OuterXml)

$cmp.XmlExportType = 3 
$cmp.XMLAsString = 0

 $oItem = $cmp.GetBusinessObject(4)
 $BPCode = ""
if ($oItem.GetByKey('TS01096_CH')) {# $oItem.SaveXML('c:\temp\t15.xml')

 [xml]$xmlItem =  $oItem.GetAsXML()
 #[xml]$xmlItem = gc 'c:\temp\t15.xml'
 $xmlItem.PreserveWhitespace = $true
 # $xmlItem.SelectNodes('//row/OrderIntervals') | %{$_.ParentNode.removechild($_)} 
 #$xmlItem.SelectNodes("//row/LeadTime")
 #$xmlItem.SelectNodes("//row/ProcurementMethod") 
# $xmlItem.SelectNodes("//row/BPCode")
# $xmlItem.SelectNodes("//row/ToleranceDays")
 $nodes = $xmlItem.SelectNodes('/BOM/BO') 
 foreach($node in $nodes) {
    $node.Items.row.LeadTime = ‘1250’
    $node.items.row.ProcurementMethod = "bom_Buy"
    $node.items.row.ToleranceDays = ‘1520’
    #IF ([string]::IsNullOrEmpty($node.InnerText)){$node.ParentNode.removechild($node)} 
    $node.AppendChild($xml)
    
    # if prefer supplier is to be blank 
    if ([string]::IsNullOrEmpty($BPCode) -and $node.ItemPreferredVendors.row.BPCode -ne $null) {$NODE.RemoveChild("ItemPreferredVendors")}
    $node.ItemPreferredVendors.row.BPCode = "306565V"
} 



#$xmlItem.Save('c:\temp\t16.xml') 

  #$oItem.Browser.ReadXml('c:\temp\t16.xml',0)
  
  $oItem.Browser.ReadXml($xmlItem,0)
  $oItem.Update() ; $cmp.GetLastErrorDescription()
}

