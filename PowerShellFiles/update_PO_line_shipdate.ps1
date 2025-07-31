#-----------connect to SAPB1-------
	 $cmp = New-Object -COMObject 'SAPbobsCOM.Company'
	$cmp.Server ="SZ-SAPTST82" 
	$cmp.CompanyDB = "SAPB1_sz_TST"
	$cmp.DbServerType = 8
	$cmp.UserName = "jieva"
	$cmp.Password ="Ves-1234"
	$cmp.DbUserName="Butterfly"
    $cmp.DbPassword="buTterF1y"
	$cmp.LicenseServer = "sz-TSTSAPLIC92"
	[VOID]$cmp.Connect()
	if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() 
}
      $cmp.XMLAsString = 1
      $cmp.XmlExportType = 3 


$xml =@'
<BOM>	<BO>		<AdmInfo>			<Object>22</Object>			<Version>2</Version>		</AdmInfo>		<Documents>			<row>				<DocEntry>$($docentry)</DocEntry>			</row>		</Documents>		<Document_Lines>			$($lines -join "`n")		</Document_Lines>	</BO></BOM>
'@

$docline=@'
<row> <LineNum>$($line.linenum)</LineNum> <ShipDate>$($line.shipdate.Replace( ".", ""))</ShipDate></row>
'@

$starttime=get-date
function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}


$GroupsData=Import-Csv  -LiteralPath "\\sz-sz-sapb1app\SZ_PO_ShipDate\SZ_PO_NEW_SHIPDATE3.csv" |Group-Object DocEntry

foreach($grp in $GroupsData) {
      $fxml=""
      $po = $cmp.GetBusinessObject(22)   
       if ($po.getbykey($grp.Name))
        { $docentry=$grp.Name
          $lines=foreach($line in $grp.Group){ $ExecutionContext.InvokeCommand.ExpandString($docline)   }
          $fxml=$ExecutionContext.InvokeCommand.ExpandString($xml) 
       
          $po.Browser.ReadXml($fxml,0)
          [void]$po.Update()
          #$cmp.GetLastErrorDescription()
          
        }


}

$endtime=get-date

[void]$cmp.Disconnect()
Release-Ref($cmp)

<#save hashtable to file
 $outTime=Get-Date -format 'yyyyMMddHms'
 $outpath = "c:\temp\$outTime.log"
 $log= $hash.Keys | Sort-Object | ForEach-Object{ $hash[$_] }   
  set-Content $outpath -Value $log -Encoding Unicode
  #>


Write-Host -ForegroundColor Red ('Total Runtime: ' + ($endtime - $starttime).TotalSeconds)





