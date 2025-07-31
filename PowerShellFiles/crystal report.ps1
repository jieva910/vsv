$cmp = New-Object -ComObject 'SAPBOBSCom.company' ;
 $UDOName = 'VES_CONSGN' ;
 $childUDOName = 'VES_CONSGNDTL' ; 
 function SendeMail($content,$sub,$to)
 { Send-MailMessage -Body $content -From SAPB1APP@Vsv.COM  -Subject $sub  -To $to   -SmtpServer 'APMailrelay.vesuvius.com' -port 25  };
  if($cmp.Connected -eq $true){[VOID]$cmp.Disconnect()}
 $cmp.SLDServer = 'sz-tstsaplic92:40000';
 $cmp.Server = 'SZ-SAPTST82';
 $cmp.CompanyDB = 'SAPB1_SZ_TST';
 $cmp.DbServerType = 8;
 $cmp.UseTrusted = $true;
 $cmp.UserName = 'zhaojac';
 $cmp.Password ='Ves-123456';
 $cmp.Connect(); 
 $cmp.GetLastErrorDescription()


# Dim $oCmpSrv As SAPbobsCOM.CompanyService
#Dim $oReportLayoutService As ReportLayoutsService
#Dim $oReportLayout As ReportLayout
#Dim $oReportLayoutParam As ReportLayoutParams

# Get report layout service
$oCmpSrv = $cmp.GetCompanyService()
$oReportLayoutService = $oCmpSrv.GetBusinessService(232)  #[SAPbobsCOM.ServiceTypes]::ReportLayoutsService



# Set parameters
$oReportLayoutParam = $oReportLayoutService.GetDataInterface(2) # [SAPbobsCOM.ReportLayoutsServiceDataInterfaces]::rlsdiReportLayoutParams
 $oReportLayoutPrintParams = $oReportLayoutService.GetDataInterface(4) # rlsdiReportLayoutPrintParams
$oReportLayoutParam.LayoutCode = "POR20007"

# Get report layout
$oReportLayout = $oReportLayoutService.GetReportLayout($oReportLayoutParam)

# Add report layout
$oReportLayoutService.AddReportLayout($oReportLayout)

# printing inside SAPB1
 $oReportLayoutPrintParams.DocEntry = 66842
 $oReportLayoutPrintParams.LayoutCode ="POR20007"
$oReportLayoutService.Print($oReportLayoutPrintParams)


# Print to PDF via crystal report component
#load the assemblies

[reflection.assembly]::LoadWithPartialName(‘System.Drawing.Printing’)

[reflection.assembly]::LoadWithPartialName(‘CrystalDecisions.Shared’)

[reflection.assembly]::LoadWithPartialName(‘CrystalDecisions.CrystalReports.Engine’)





#setup Crystal Document object ############################################################################
$ReportLocation = "C:\temp\test\Purchase Order - Vesuvius 2017.rpt"

$ReportDestination = "C:\temp\test\reporttest.pdf"

$Report = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument

#load the Crystal Report

$Report.Load($ReportLocation)

$Report.SetParameterValue("DocKey@",66414)

#if your report requires a database connection. Otherwise you can delete this line.

# $Report.SetDatabaseLogon(‘username’,’password’)

#run report and export to pdf

$Report.ExportToDisk([CrystalDecisions.Shared.ExportFormatType]::PortableDocFormat,'c:\temp\66414.pdf')

$Report.Close()




#create the required printer and page settings ############################################################

$PrinterName = “printer0001”

$PrintOptions = New-Object System.Drawing.Printing.PrinterSettings

$PageOptions = New-Object System.Drawing.Printing.PageSettings

#add the desired printer or it will print to the default printer

$PrintOptions.PrinterName = $PrinterName

$PrintOptions.Copies = 1

#print

$Report.PrintToPrinter($PrintOptions, $PageOptions, $false)

$Report.Dispose()