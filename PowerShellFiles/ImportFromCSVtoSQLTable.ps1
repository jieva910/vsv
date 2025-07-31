#------------------------初始化区------------------------------------ 
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") #加载WinForm库
$oCompany = new-object -ComObject "SAPBOBSCOM.Company"
$ServerInstance = ""
$DatabaseName = ""
# Connect to SBO via UI API
Function SetApplication {
              $SboGuiApi = New-Object -comobject "SAPbouiCOM.SboGuiApi"
              $sConnectionString =  "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
              $SboGuiApi.Connect($sConnectionString)
              $SboGuiApi.GetApplication()
          
        }
 #Connect with connection string
$SBO_Application = SetApplication
  function SetConnectionContext {
             $sCookie = $oCompany.GetContextCookie()
             $sConnectionContext = $SBO_Application.Company.GetConnectionContext($sCookie)
             If ($oCompany.Connected ){$oCompany.Disconnect()}
             return $oCompany.SetSboLoginContext($sConnectionContext)
        }
    # Connect to SBO via DI API
    Function ConnectToCompany {
       Return $oCompany.Connect()
    }

function UI_DI_Conn
{
    # connect to DI 
    try {
        if (SetConnectionContext -ne 0 ) {msgshow 'Connection to SAPB1' "Failed setting a connection to DI API" ; RETURN}
        if (ConnectToCompany -ne 0 ) {msgshow 'Connection to SAPB1' "Failed connecting to the company's Data Base" ; RETURN}
     # Write-Host -ForegroundColor Cyan "DI Connected To: " $oCompany.LicenseServer $oCompany.CompanyName
     msgshow  'Connection to SAPB1' "connected to the $($oCompany.companydb) company's Data Base"
 
    }
      
    catch{
    msgshow 'Excpetion' $_.exception.message
    }
    finally
    {
      $script:ServerInstance = $oCompany.Server
      $script:DatabaseName = $oCompany.companydb
	  $Choose12.text  = $oCompany.companydb 
    }
}




function Select-File ($InitialDirectory) {
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Please Select File"
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.filter = "CSV (*.CSV)| *.CSV"
    If ($OpenFileDialog.ShowDialog() -eq "Cancel") {
        [System.Windows.Forms.MessageBox]::Show("No File Selected. Please select a file !", "Error", 0, 
        [System.Windows.Forms.MessageBoxIcon]::Exclamation)
        $result = $null
    } 
    else { 
        $result = $OpenFileDialog.FileName
    }
    $OpenFileDialog.Dispose()

    return $result
}   

function msgshow($msgtitle,$msgbody)  {
	Add-Type -AssemblyName PresentationCore,PresentationFramework 
$ButtonType = [System.Windows.MessageBoxButton]::OK
 $MessageIcon = [System.Windows.MessageBoxImage]::Information
  $MessageBody = $msgbody
  $MessageTitle = $msgtitle

 [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle)
}

FUNCTION GenerateAPDraft
{
  try {
   $sql_DocEntry  = "select * from dbo.VES_GRNS where isnull(Status,'')<>'C'"

      $oRs_sp = $oCompany.getbusinessobject(300)
       $oRs = $oCompany.getbusinessobject(300)
      $oRs_sp.doquery("exec dbo.sp_VES_SDInvoice_exec 0")
      $oRs.doquery($sql_DocEntry)
      if ($oRs.Eof){msgshow 'Information' 'No Record matching,please check';RETURN}
      $xml =[xml]$oRs.getasxml()
      $nodes = $xml.selectnodes("//row")
      $groupbyDocEntry = $nodes|  Group-Object -Property Cardcode,SDNumber,SDInvDate,SDInvTotal
    
     foreach($grps in $groupbyDocEntry)
     { 
     
        $oCompany.StartTransaction() 
         $sdnum =''
         $msg=''
         $oOPCHDrafts = $oCompany.GetBusinessObject(112)
		  $oOPCHDrafts.DocObjectCode = 18
		  $oOPCHDrafts.HandWritten = 0
	     $oOPCHDrafts.CardCode = ($grps.Name -split ",")[0]  
	     $oOPCHDrafts.DocDate = get-date -Format "yyyy-MM-dd" 
        #  $oOPCHDrafts.DocTotal =  ($grps.Name -split ",")[3] 
          $dateInt = [int](($grps.Name -split ",")[2])  
          $oOPCHDrafts.TaxDate = "{0:####-##-##}" -f  $dateInt 
          $sdnum = ($grps.Name -split ",")[1] 
          $oOPCHDrafts.NumAtCard = $sdnum
          foreach($grp in $grps.Group)
          {  
            $i = 0 
            $oOPCHDrafts.lines.basetype = 20 
	        $oOPCHDrafts.lines.baseentry = $grp.GRNEntry 
	        $oOPCHDrafts.lines.baseline =$grp.linenum 
			$oOPCHDrafts.lines.UnitPrice =$grp.Price 
	        $oOPCHDrafts.lines.Quantity =$grp.GRNOpenQty  
            $oOPCHDrafts.Lines.TaxTotal = [double]$grp.vatsum2
	        $oOPCHDrafts.Lines.SetCurrentLine($i) 
	        $oOPCHDrafts.lines.add()
 
            $i++
           }
          $rt = $oOPCHDrafts.add()
           if($rt -eq 0 )
            { $oRs_updatesql=$oCompany.GetBusinessObject(300)
            $oRs_updatesql.doquery("update dbo.VES_GRNS set Status ='C',updatedate=current_timestamp where  SDNumber = $sdnum")
          #  $oRs_updatesql.doquery("update dbo.VES_JXI0 set Docstatus = 'C',Comments='',updatedate = getdate()  WHERE  isnull(Docstatus,'')<>'C' and  SDNumber =$($sdnum)")
             $oRs_updatesql.doquery("update I1 set i1.[status] ='C',updatedate=current_timestamp  FROM dbo.[VES_JXI0] I1 WITH(NOLOCK) WHERE I1.SDNumber =$($sdnum)")
            #$msg = "$($sdnum) has been finished Auto APINVOICE in $($oCompany.CompanyDB)" 
             }
           else { 
            $oRs_updatesql.doquery("update dbo.VES_GRNS set Status ='F',updatedate=current_timestamp where  SDNumber = $sdnum")
            $oRs_updatesql.doquery("insert into VES_JXLO (SDNumber,ErrorCode,ErrorMessage) values ($($sdnum),$([string]$oCompany.GetLastErrorCode()),$($oCompany.GetLastErrorDescription())")
             $oRs_updatesql.doquery("update I1 set i1.[status] ='F',updatedate=current_timestamp  FROM  dbo.[VES_JXI0] I1  WHERE I1.SDNumber =$($sdnum)")
             #$msg = "$($sdnum) failed to auto APINVOICE in $($oCompany.CompanyDB) ,$($failreason)"
           }
           $oCompany.EndTransaction([SAPbobsCOM.BoWfTransOpt]::wf_Commit)
        }
	 msgshow 'Confirmed' 'Finished,please check AP Invoice draft reports'	
    }
        catch [exception]
        {  msgshow 'sapb1 GenerateAPDraft exception' $_.Exception.Message ;RETURN  }
        
         
      	
	
}
function ImportFromCSVtoSQL($sourceFile)
{
 
 try {
       $csv = import-csv  $sourceFile -Delimiter ","  #'C:\dell\CS VAT list2.csv'
# group by sdnum
$Grps = $csv | Group-Object -Property 发票号码

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
msgshow "Confirm" 'Data importing from CSV to SQL database finished'
    }
 catch{
    [System.Windows.MessageBox]::Show($_.exception.message,'Excpetion','Ok','Error')
    }
   
}
###############
# Main routine
###############

[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.AutoSize                   = $true
$Form.text                       = "ImportData to SapB1"
$Form.TopMost                    = $true
#----------------------

$Choose1                      = New-Object system.Windows.Forms.Label
$Choose1.text                 = ""
$Choose1.AutoSize             = $true
$Choose1.width                = 25
$Choose1.height               = 10
$Choose1.location             = New-Object System.Drawing.Point(28,20)
$Choose1.ForeColor            = "#000000"
$Choose11                      = New-Object system.Windows.Forms.Label
$Choose11.text                 = "Connected to : "
$Choose11.AutoSize             = $true
$Choose11.width                = 25
$Choose11.height               = 10
$Choose11.location             = New-Object System.Drawing.Point(28,80)
$Choose11.ForeColor            = "#093c76"

$Choose12                      = New-Object system.Windows.Forms.Label
$Choose12.text                 = ""
$Choose12.AutoSize             = $true
$Choose12.width                = 250
$Choose12.height               = 30
$Choose12.location             = New-Object System.Drawing.Point(120,80)
$Choose12.ForeColor            = "#FF0000"

$Sel                        = New-Object system.Windows.Forms.TextBox
$Sel.AutoSize               = $true
$Sel.width                  = 250
$Sel.height                 = 30
$Sel.location               = New-Object System.Drawing.Point(120,40)
$Sel.Text                   = "Selected"

$Choose2                        = New-Object System.Windows.Forms.Button
$Choose2.text                   = "Selected File"
$Choose2.AutoSize               = $true
$Choose2.width                  = 90
$Choose2.height                 = 20
$Choose2.location               = New-Object System.Drawing.Point(28,38)
$Choose2.ForeColor              = "#ffffff"
$Choose2.BackColor              = "#093c76"
$Choose2.Add_Click({ $Sel.Text = Select-File })

$ConnecToSAPB1                         = New-Object system.Windows.Forms.Button
$ConnecToSAPB1.BackColor               = "#6996c8"
$ConnecToSAPB1.text                    = "ConnecToSAPB1"
$ConnecToSAPB1.width                   = 120
$ConnecToSAPB1.height                  = 30
$ConnecToSAPB1.location                = New-Object System.Drawing.Point(80,190)
$ConnecToSAPB1.Add_Click({UI_DI_Conn })

$ImportData                         = New-Object system.Windows.Forms.Button
$ImportData.BackColor               = "#6996c8"
$ImportData.text                    = "ImportData"
$ImportData.width                   = 90
$ImportData.height                  = 30
$ImportData.location                = New-Object System.Drawing.Point(210,190)
$ImportData.Add_Click({  ImportFromCSVtoSQL $Sel.Text $(Split-Path $Sel.Text)  })

$GenerateAPDraft                         = New-Object system.Windows.Forms.Button
$GenerateAPDraft.BackColor               = "#6996c8"
$GenerateAPDraft.text                    = "GenerateAPDraft"
$GenerateAPDraft.width                   = 120
$GenerateAPDraft.height                  = 30
$GenerateAPDraft.location                = New-Object System.Drawing.Point(310,190)
$GenerateAPDraft.Add_Click({GenerateAPDraft})


$Close                         = New-Object system.Windows.Forms.Button
$Close.BackColor               = "#6996c8"
$Close.text                    = "Close"
$Close.width                   = 98
$Close.height                  = 30
$Close.location                = New-Object System.Drawing.Point(450,190)
$Close.Add_Click({$Form.Close()})

#----------
#

$Form.Controls.AddRange(@($Choose1,$Choose11,$Choose12,$Sel, $Choose2,$ImportData,$ConnecToSAPB1,$GenerateAPDraft,$Close))
[void] $Form.ShowDialog()


# clean up the form
$Form.Dispose()

