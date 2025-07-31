$cmp  = new-object -ComObject "sapbobscom.company"

$cmp.Server = "SZ-SAP01"
$cmp.SLDServer ="SZ-SAPLIC92:40000"
$cmp.CompanyDB = "SAPB1_YK"
$cmp.DbServerType = 8
$CMP.UseTrusted = 1 
$cmp.UserName="\"
$cmp.Password=""

$cmp.Connect()
$cmp.GetLastErrorDescription()


 
# Get report layout service
$oCmpSrv = $cmp.GetCompanyService()
      $oReportLayoutService = $oCmpSrv.GetBusinessService(232);
       $oPrintParam = $oReportLayoutService.GetDataInterface(4);
       $oPrintParam.LayoutCode = "INV20009";
$csv = Import-Csv C:\Dell\ARInvoiceYK.csv
foreach($r in $csv)
{


       $oPrintParam.DocEntry = $r.DocEntry;
       $oReportLayoutService.Print($oPrintParam);
$rtTrue =$false
do 
{
  Start-Sleep -s 15
    $filename = $r.DocNum
$wshell = new-object -com wscript.shell
$rtTrue =  $ture
if($rtTrue){$wshell.sendkeys("$($filename)");$wshell.sendkeys("{Enter}");break}
  
} until  ($rtTrue -eq $ture)


}





# alternative method 

[void][reflection.assembly]::LoadWithPartialName(‘System.Drawing.Printing’)
[void][reflection.assembly]::LoadWithPartialName(‘CrystalDecisions.Shared’)
[void][reflection.assembly]::LoadWithPartialName(‘CrystalDecisions.CrystalReports.Engine’)

#setup Crystal Document object ############################################################################
$ReportLocation = "D:\ARINVOICE\Invoices_YK_2025.rpt"
$Report = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
$CSV  = IMPORT-CSV D:\ARINVOICE\ARInvoiceYK.csv -Delimiter ","
FOREACH($R IN $CSV)
{
 $Report.Load($ReportLocation)
  $Report.SetParameterValue("DocKey@",$r.DocEntry)

#if your report requires a database connection. Otherwise you can delete this line.
 $Report.SetDatabaseLogon(‘CRDeveloper’,’CRDeveloper’)
 $ReportDestination = "D:\tstpdf\$($R.DocNum).pdf"
#run report and export to pdf
$Report.ExportToDisk([CrystalDecisions.Shared.ExportFormatType]::PortableDocFormat,$ReportDestination)
$Report.Close()

Send-MailMessage -Body 'Please find attached report for ICRP Invoice @YK.' -From sbp.crystal@vesuvius.com  -Subject "ICRP@YK@$($R.DocNum)@" -To 'Live.ICRP@vesuvius.com' -Cc 'evan.ji@vesuvius.com' -Attachments $ReportDestination  -SmtpServer 'APMailrelay.vesuvius.com' -port 25
Write-Host "PDF Report generated at: $ReportDestination"

}


