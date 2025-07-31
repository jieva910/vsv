$csv = import-csv 'C:\dell\CS VAT list2.csv'
# group by sdnum
$Grps = $csv | Group-Object -Property 发票号码


$ServerInstance = "sz-sapstg91"
$DatabaseName = "sapb1_cs_tst"
$ConnString = "Server=$ServerInstance;Database=$DatabaseName;Integrated Security=SSPI;"
$conn = New-Object System.Data.SqlClient.SqlConnection $ConnString
$conn.Open()


foreach($grp in $Grps)
{ 
  
  $datatable = [System.Data.DataTable]::new()
[void]($datatable.Columns.Add("SDNumber","System.String"))
[void]($datatable.Columns.Add("DocDate","System.String"))
[void]($datatable.Columns.Add("SupplierName","System.String"))
[void]($datatable.Columns.Add("Itemcode","System.String"))
[void]($datatable.Columns.Add("Specification","System.String"))
[void]($datatable.Columns.Add("UnitMsr","System.String"))
[void]($datatable.Columns.Add("Qty","System.String"))
[void]($datatable.Columns.Add("Price","System.String"))
[void]($datatable.Columns.Add("LineTotal","System.String"))
[void]($datatable.Columns.Add("VatPrct","System.String"))
[void]($datatable.Columns.Add("VatSum","System.String"))
[void]($datatable.Columns.Add("LinkedPO","System.String"))

 foreach($r in $grp.Group)
{  
   
   $match = ((($r.发票备注 -replace '（.*）', '' | sls -Pattern '12\d+' -AllMatches).Matches | ? {$_.Value.Length -eq 10}).Value | select -Unique) -join ”,"
   $itemcode = [regex]::Match($r.'货物或应税劳务、服务名称',"[a-zA-Z]\w.*") | Select-Object -ExpandProperty Value
   if ([string]::IsNullOrEmpty($itemcode)) {$itemcode = $r.'货物或应税劳务、服务名称'}
   [void]($dataTable.Rows.Add($r.开票日期,$r.销方名称,$r.发票号码,$r.'货物或应税劳务、服务名称',	$r.规格型号,$r.单价,$r.单位,$r.数量,$r.金额,	[int]($r.税率.Replace('%','')),$r.税额, $match))
}

$query = "dbo.sp_VES_SDInvoice_ImportFromCsv"
$cmd = New-Object System.Data.SqlClient.SqlCommand
$cmd.Connection = $conn
$cmd.CommandType = [System.Data.CommandType]"StoredProcedure"
$cmd.CommandText = $Query
$cmd.Parameters.Add("@SDnum", [System.Data.SqlDbType]::NVarChar) | out-null
$cmd.Parameters["@SDnum"].Value =[string]($grp.Name)
$cmd.Parameters.Add("@SDInvLine_insert", [System.Data.SqlDbType]::Structured) | Out-Null
$cmd.Parameters["@SDInvLine_insert"].Value = $dataTable

$cmd.ExecuteNonQuery() | Out-Null

}


$conn.Close()