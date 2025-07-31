


#$Serv="SZ-SAP01"
#$Database ="SAPB1-SZ"
$Quit =0
$SqlConn = New-Object System.Data.SqlClient.SqlConnection
#连接 MSSQL
   $SqlConn.ConnectionString = "Data Source=$ServInitial Catalog=$Databaseuid='Butterfly'password='buTterF1y'"

   #打开数据库连接
   $SqlConn.open()
   $SqlCmd = New-Object System.Data.SqlClient.SqlCommand('select * from oadm',$SqlConn)
   $SqlCmd.CommandTimeout = 0
  $datareader= $SqlCmd.ExecuteReader() 
  if (!$datareader.HasRows) {   $datareader.Close()exit}
 
 
Add-Type -Path 'C:\Program Files (x86)\SAP\SAP Business One DI API\DI API 90\SAPbobsCOM90.dll'
[System.Reflection.Assembly]::LoadWithPartialName('SAPbobsCOM')

$cmp  = new-object -ComObject "sapbobscom.company"

$cmp.Server = "SZ-SAPSTG91"
$cmp.SLDServer ="S"
$cmp.CompanyDB = "SAPB1_CS_TST"
$cmp.DbServerType =8
$cmp.DbUserName ="butterfly"
$cmp.DbPassword=""
$cmp.UserName="montova"
$cmp.Password=""
$cmp.UseTrusted = 0 
$cmp.Connect()
$cmp.GetLastErrorDescription()


$cmp.XmlExportType  = 3 

$udo = 'NominationClient'
# export UDO to xml
$oCompServic = $cmp.GetCompanyService()
$oGeneralServic = $oCompServic.GetGeneralService($udo)
 $oGeneralData = $oGeneralServic.GetDataInterface(1)
$oGeneralParams  = $oGeneralServic.GetDataInterface([SAPbobsCOM.GeneralServiceDataInterfaces]::gsGeneralDataParams)  
  $oGeneralParams.setProperty("Code","00000005")
  $oGeneralData = $oGeneralServic.GetByParams($oGeneralParams)

   $oGeneralData.ToXMLFile('c:\temp\x.xml')


   $oitem = $cmp.GetBusinessObject(4)

   $oitem.GetByKey('ZZ17261TST01')

   $oitem.SaveXML('c:\temp\item.xml')



 ######## #  tst auto ap invoice  2024.05.10  
     # update 2024.8.15
      $sql_DocEntry  = "select * from dbo.VES_GRNS where isnull(Status,'')<>'C'"

      $oRs_sp = $cmp.getbusinessobject(300)
       $oRs = $cmp.getbusinessobject(300)
      $oRs_sp.doquery("exec dbo.sp_VES_SDInvoice_exec")
      $oRs.doquery($sql_DocEntry)
      if ($oRs.Eof){break}
      $xml =[xml]$oRs.getasxml()
      $nodes = $xml.selectnodes("//row")
      $groupbyDocEntry = $nodes|  Group-Object -Property Cardcode,SDNumber,SDInvDate,SDInvTotal

     foreach($grps in $groupbyDocEntry)
     { 
      try {
        $cmp.StartTransaction() 
         $sdnum =''
         $msg=''
         $oOPCH = $cmp.GetBusinessObject(18)
	     $oOPCH.CardCode = ($grps.Name -split ",")[0]  
	     $oOPCH.DocDate = get-date -Format "yyyy-MM-dd" 
        #  $oOPCH.DocTotal =  ($grps.Name -split ",")[3] 
          $dateInt = [int](($grps.Name -split ",")[2])  
          $oOPCH.TaxDate = "{0:####-##-##}" -f  $dateInt 
          $sdnum = ($grps.Name -split ",")[1] 
          $oOPCH.NumAtCard = $sdnum
          foreach($grp in $grps.Group)
          {  
            $i = 0 
            $oOPCH.lines.basetype = 20 
	        $oOPCH.lines.baseentry = $grp.GRNEntry 
	        $oOPCH.lines.baseline =$grp.linenum   
	        $oOPCH.lines.Quantity =$grp.GRNOpenQty  
            $oOPCH.Lines.TaxTotal = [double]$grp.vatsum2
	        $oOPCH.Lines.SetCurrentLine($i) 
	        $oOPCH.lines.add()
 
            $i++
           }
          $rt = $oOPCH.add()
           if($rt -eq 0 )
            { $oRs_updatesql=$cmp.GetBusinessObject(300)
            $oRs_updatesql.doquery("update dbo.VES_GRNS set Status ='C',updatedate=current_timestamp where  SDNumber = $sdnum")
          #  $oRs_updatesql.doquery("update dbo.VES_JXI0 set Docstatus = 'C',Comments='',updatedate = getdate()  WHERE  isnull(Docstatus,'')<>'C' and  SDNumber =$($sdnum)")
             $oRs_updatesql.doquery("update I1 set i1.[Linestatus] ='C',updatedate=current_timestamp  FROM dbo.[VES_JXI1] I1 WITH(NOLOCK) WHERE I1.SDNumber =$($sdnum)")
            $msg = "$($sdnum) has been finished Auto APINVOICE in $($cmp.CompanyDB)" 
             }
           else { 
            $oRs_updatesql.doquery("update dbo.VES_GRNS set Status ='F',updatedate=current_timestamp where  SDNumber = $sdnum")
            $oRs_updatesql.doquery("insert into VES_JXLO (LogDate,SDNumber,ErrorCode,ErrorMessage) values (CURRENT_TIMESTAMP,$($sdnum),$([string]$cmp.GetLastErrorCode()),$($cmp.GetLastErrorDescription())")
             $oRs_updatesql.doquery("update I1 set i1.[Linestatus] ='F',updatedate=current_timestamp  FROM  dbo.[VES_JXI1] I1  WHERE I1.SDNumber =$($sdnum)")
             $msg = "$($sdnum) failed to auto APINVOICE in $($cmp.CompanyDB) ,$($failreason)"
           }
           $cmp.EndTransaction([SAPbobsCOM.BoWfTransOpt]::wf_Commit)
        }
        catch [exception]
        {  $msg =$_   }
        finally
        { 
        $oRs_updatesql.doquery("insert into VES_JXLO (LogDate,SDNumber,ErrorCode,ErrorMessage) values (CURRENT_TIMESTAMP,$($sdnum),1,$($msg)")
        #Send-MailMessage -Body $msg -From AutoAPInvoice@SAPB1APP.COM  -Subject 'SAPB1 AutoAPInvoice' -To 'evan.ji@vesuvius.com'  -SmtpServer 'APMailrelay.vesuvius.com' -port 25
        }
         
      }


     
     # TEST Outgoing payment doc  2024.05.29

   # $oIPD = $cmp.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oPaymentsDrafts)  Drafts
   $oIPD = $cmp.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oVendorPayments)
$oIPD.DocType = [SAPbobsCOM.BoRcptTypes]::rSupplier
$oIPD.CardCode = "307246V"
$oIPD.DocDate = Get-Date
$oIPD.TransferAccount = "221100-14"
$oIPD.TransferSum = [decimal]61800
$oIPD.TransferDate = Get-Date
$oIPD.TransferReference = "test pay from di"
$oIPD.DocObjectCode = [SAPbobsCOM.BoPaymentsObjectType]::bopot_OutgoingPayments
$oIPD.Invoices.InvoiceType = [SAPbobsCOM.BoRcptInvTypes]::it_PurchaseDownPayment
$oIPD.Invoices.DocEntry = 1199
$oIPD.Invoices.SumApplied = [decimal]61800
Write-Host  $oIPD.Add() $cmp.GetLastErrorDescription()
      

# test JE for internal reconcilation 2024.5.30

$oJE = $cmp.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oJournalEntries)
$oJE.ReferenceDate =Get-Date
$oJE.DueDate=Get-Date
$oJE.Memo = 'Outgoing Payments-307246V-SZ-001'
$oJE.Lines.AccountCode='221100-14'
$oJE.Lines.Debit=[decimal]61800
$oJE.Lines.Add()
$OJE.Lines.AccountCode ='172000-14-01-06'
$oJE.Lines.Credit=[decimal]61800
Write-Host $oJE.ADD()  $cmp.GetLastErrorDescription()

# TEST document draft 


    #close  the Documents object
    $vDrafts=$cmp.GetBusinessObject(112)
    $ors = $cmp.GetBusinessObject(300)
    $ors.doquery("SELECT t.DocEntry FROM ODRF t WHERE t.DocStatus = 'o'  AND t.ObjType = 18")
     if ($oRs.Eof){break}
      $xml =[xml]$oRs.getasxml()
      $nodes = $xml.selectnodes("//row")
      foreach($n in $nodes)
      {
        if ( $vDrafts.getbykey($n.Docentry)){
          $vDrafts.Remove(), $cmp.GetLastErrorDescription()
        }
      }

 





   $oOPCH = $cmp.GetBusinessObject(18)
	     $oOPCH.CardCode = '301082V'
	     $oOPCH.DocDate = get-date -Format "yyyy-MM-dd" 
          #$oOPCH.DocTotal =  71.7       
          $oOPCH.TaxDate = get-date -Format "yyyy-MM-dd"  
         
          $oOPCH.NumAtCard = 'test3 vatsum'
         
            $oOPCH.lines.basetype = 20 
	        $oOPCH.lines.baseentry = 81158	
	        $oOPCH.lines.baseline =22 
	        $oOPCH.lines.Quantity =50  
           # $oOPCH.Lines.TaxLiable = 0 
            $oOPCH.Lines.TaxTotal = [double]8.21
	     	        $oOPCH.lines.add()
          $rt = $oOPCH.add()

          $cmp.GetLastErrorDescription()
		  


## 2024-12-30  orct 科目收款

 
   $oIPD = $cmp.GetBusinessObject(24)
   $oIPD.DocObjectCode = [SAPbobsCOM.BoPaymentsObjectType]::bopot_IncomingPayments
$oIPD.DocType = [SAPbobsCOM.BoRcptTypes]::rAccount
$oIPD.DocDate = Get-Date
$oIPD.CardCode = "173000-02"
$oIPD.DocCurrency ='RMB'
$oIPD.JournalRemarks ='TEST INCOMING PAY'
$oIPD.LocalCurrency = 1
$oIPD.TaxDate= Get-Date
$oIPD.TransferAccount = "570140-01"
   $oIPD.TransferDate =  Get-Date
   $oIPD.TransferSum =120
$oIPD.TransferReference = "test INCMONG pay from di"
$oIPD.AccountPayments.AccountCode="173000-02"
$oIPD.AccountPayments.SumPaid = 120
$oIPD.AccountPayments.GrossAmount=120

Write-Host  $oIPD.Add() $cmp.GetLastErrorDescription()



declare @Rt int ,@msg NVARCHAR(255)
exec [dbo].[usp_UnlockUser2]
    @USER ='zhangrac',
	@CompanyDB= 'sapb1_cs_tst',
	@DILIC ='sz-tstsaplic92',
	@DISVR='sz-sapstg91',
	@DBTYPE =10 ,
	@DIUSER ='jieva',
	@DIPWD ='Ves-123456',
	@DBID ='butterfly',
	@DBPWD ='buTterF1y',
	--@DITURST BIT = 0 ,
  --  @newPassword VARCHAR(100) = 'Ves?147369',  -- New parameter for new password when resetting
  --  @FunctionName VARCHAR(50) = 'userunlock',     -- userunlock ,setpassword ,COAUnblock，COAblock
	 @AccountCode = '141600-06',
	-- @isacctunblock BIT = 1 ,
	@Result =@rt OUTPUT ,
	@errorDescription =@msg OUTPUT
;

select @msg


# remove bill to of cardcode

$obp = $cmp.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oBusinessPartners)
if ($obp.GetByKey('443395')) {
    $cardCode = $obp.CardCode
    $billToIndices = @()
    $nonMatchingIndices = @()

    # First pass: Identify all Bill-to addresses
    for ($i = 0; $i -lt $obp.Addresses.Count; $i++) {
        $obp.Addresses.SetCurrentLine($i)
        if ($obp.Addresses.AddressType -eq [SAPbobsCOM.BoAddressType]::bo_BillTo) {
            $billToIndices += $i
            if ($obp.Addresses.AddressName.Trim() -ne $cardCode) {
                $nonMatchingIndices += $i
            }
        }
    }

    $billToCount = $billToIndices.Count
    $madeChanges = $false

    # Case 1: Single Bill-to address that doesn't match CardCode
    if ($billToCount -eq 1 -and $nonMatchingIndices.Count -eq 1) {
        $obp.Addresses.SetCurrentLine($nonMatchingIndices[0])
        $obp.Addresses.AddressName = $cardCode
        $madeChanges = $true
        Write-Host "Updated single Bill-to address to match CardCode"
    }
    # Case 2: Multiple Bill-to addresses
    elseif ($billToCount -ge 2) {
        # Remove non-matching addresses in reverse order (to avoid index shifting)
        $nonMatchingIndicesSorted = $nonMatchingIndices | Sort-Object -Descending
        
        foreach ($index in $nonMatchingIndicesSorted) {
            $obp.Addresses.SetCurrentLine($index)
            $obp.Addresses.Delete()
            $madeChanges = $true
            Write-Host "Removed non-matching Bill-to address at index $index"
        }
    }

    # Save changes if modifications were made
    if ($madeChanges) {
        $updateResult = $obp.Update()
        if ($updateResult -ne 0) {
            $errorMsg = $cmp.GetLastErrorDescription()
            Write-Host "Update failed: $errorMsg"
        } else {
            Write-Host "Customer $cardCode updated successfully"
        }
    } else {
        Write-Host "No changes needed for customer $cardCode"
    }
} else {
    Write-Host "Customer 159744 not found"
}
