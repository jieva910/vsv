# Update shipate and send mail with HTML body


#-----------connect to SAPB1-------
[System.Reflection.Assembly]::LoadWithPartialName('SAPbobsCOM')
 $cmp = New-Object -COMObject SAPbobsCOM.Company
	$cmp.Server ="SZ-SAPTST82" 
	$cmp.CompanyDB = "SAPB1_sz_TST"
	$cmp.DbServerType = 8
	$cmp.UserName = "jieva"
	$cmp.Password ="Ves-1234"
	$cmp.DbUserName="Butterfly"
    $cmp.DbPassword="buTterF1y"
	$cmp.LicenseServer = "sz-TSTSAPLIC92"
$rtval=$cmp.Connect()
	if($rtval -ne 0)
{   
	Send-MailMessage -Body $cmp.GetLastErrorDescription() -From SZ-SZ-SAPB1APP@V.COM  -To Evan.ji@vesuvius.com -SmtpServer "APMailrelay.vesuvius.com" -port 25 -Subject "Company Connection Failure:"
    exit 
}

$HTML='<h1 align="CENTER" style="color: #4485b8;">Job Finished</h1>
<h4 align="CENTER">Finished Date Time:$endtime</h4>
<table  style="margin: auto; box-shadow: 3px 3px 10px #000;" border="1">
<tbody align="CENTER">
<tr style="border-top: 2px solid #555;">
<td>DocEntry</td>
<td>PO Number</td>
<td>ErrorCode</td>
<td>Description</td>
</tr>
$($trs)
</tbody>
</table>'

$starttime=get-date

$po=$cmp.GetBusinessObject(22)
$GroupsData=Import-Csv  -LiteralPath "\\sz-sz-sapb1app\SZ_PO_ShipDate\SZ_PO_NEW_SHIPDATE3.csv" |Group-Object DocEntry
$GroupsData|ForEach-Object{  
	if ($po.getbykey($_.Name))
	{
		foreach($line in $_.Group)
		{
			$po.Lines.SetCurrentLine($line.visorder)
             $shipdate=$line.NewShipDate
             $po.Lines.ShipDate=$shipdate.Replace( ".", "-")
		}
	    $err=$po.Update()
         $docentry=$_.Name
         $PONumber=$po.DocNum
    
         $desc=$cmp.GetLastErrorDescription()
	    $trs+="<tr><td>$docentry</td><td>$PONumber</td><td>$err</td><td>$desc</td></tr>"
	}
}
$endtime=get-date
$fHTML=$ExecutionContext.InvokeCommand.ExpandString($HTML)  

Write-Host -ForegroundColor Red ('Total Runtime:'  + ($endtime - $starttime).TotalSeconds)

[void]$cmp.Disconnect()
#Set-Content -Path "c:\temp\1.html" -Value $fHTML

if($rtval -eq 0 ) {Send-MailMessage -BodyAsHtml -Body $fHTML -From SZ-SZ-SAPB1APP@V.COM  -To Evan.ji@vesuvius.com -SmtpServer "APMailrelay.vesuvius.com" -port 25 -Subject "Update Po line shipdate"}
