$cmp = new-object -ComObject "sapbobscom.company"

. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

fn_connectsapb1 $cmp 'by'

$csv = Import-Csv 'C:\Temp\AP_Invoice_category.csv' -Delimiter ","

$oudf = $cmp.GetBusinessObject(152)
$oudf.TableName = "opor"
$oudf.Name = 'Ves_Invoice_Category'
$oudf.Description = 'AP Invoice Category'
 $oudf.Type = 0  # alphanumeric
$oudf.EditSize = 40


$i = 0  # for set current line
foreach($r in $csv)
{
  
  $oudf.ValidValues.Add()
  $oudf.ValidValues.SetCurrentLine($i)
  $oudf.ValidValues.Value = $r.sort
  $oudf.ValidValues.Description = $r.category

  $i++
}
# $oudf.DefaultValue = '1'

Write-Host "Ves_Invoice_Category has been added with error $($oudf.Add()) $($cmp.GetLastErrorDescription())"


