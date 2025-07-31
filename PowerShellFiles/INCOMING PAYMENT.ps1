

$cmp  = new-object -ComObject "sapbobscom.company"

$cmp.Server = "SZ-SAP01"
$cmp.SLDServer ="SZ-SAPLIC92:40000"
$cmp.CompanyDB = "SAPB1-SZ"
$cmp.DbServerType = 8
$cmp.UseTrusted =1
$cmp.UserName="montova"
$cmp.Password="ButterSZ"

$cmp.Connect()


   $oIPD = $cmp.GetBusinessObject(24)
   $oIPD.DocObjectCode = [SAPbobsCOM.BoPaymentsObjectType]::bopot_IncomingPayments
$oIPD.DocType = [SAPbobsCOM.BoRcptTypes]::rAccount

$csv = Import-Csv C:\Temp\incoming2.csv 

foreach($r in $csv)
{
$oIPD.DocDate =  Get-Date
$oIPD.CardCode = $r.credit
$oIPD.DocCurrency ='RMB'
$oIPD.JournalRemarks =[string]$r.memo
$oIPD.LocalCurrency = 1
$oIPD.TaxDate= [datetime]$r.date
$oIPD.TransferAccount = $r.debit
   $oIPD.TransferDate =  Get-Date
   $oIPD.TransferSum =$r.amount

$oIPD.AccountPayments.AccountCode=$r.credit
$oIPD.AccountPayments.SumPaid =$r.amount
$oIPD.AccountPayments.GrossAmount=$r.amount

Write-Host  $oIPD.Add() $cmp.GetLastErrorDescription()
}



# CANCEL INCOMING PAYMENT
   $oIPD = $cmp.GetBusinessObject(24)

   $CSV_CANCEL=Import-Csv C:\Temp\incom_cancel.csv

   FOREACH($R IN $CSV_CANCEL)
   {
     $R.docnum
     $oIPD.GetByKey($R.docnum)
     $oIPD.Cancel()
     $cmp.GetLastErrorDescription()
   }