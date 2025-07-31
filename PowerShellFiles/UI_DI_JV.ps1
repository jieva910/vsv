function UI_DI_Conn
{
 
    $oCompany = New-Object -COMobject "SAPbobsCOM.Company"

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

    # connect to DI 
   
        if (SetConnectionContext -ne 0 ) {$SBO_Application.MessageBox("Failed setting a connection to DI API");break}
        if (ConnectToCompany -ne 0 ) {$SBO_Application.MessageBox("Failed connecting to the company's Data Base") ; break}
    Write-Host -ForegroundColor Cyan "DI Connected To: " $oCompany.LicenseServer $oCompany.CompanyName
}

UI_DI_Conn

$oJV = $oCompany.GetBusinessObject(28)  # JV


$oJV.JournalEntries.Memo = '备注信息'
$oJV.JournalEntries.Lines.AccountCode = '581100-01'
$oJV.JournalEntries.Lines.Debit = 100
$oJV.JournalEntries.Lines.Add()         # 新增单据行
$oJV.JournalEntries.Lines.AccountCode='252000-01'
$oJV.JournalEntries.Lines.Credit = 100

$ojv.JournalEntries.Add() |Out-Null  # 一张 JV 添加多行单据

$oJV.JournalEntries.Memo = '备注信息2'
$oJV.JournalEntries.Lines.AccountCode = '581100-01'
$oJV.JournalEntries.Lines.Debit = 200
$oJV.JournalEntries.Lines.Add()         # 新增单据行
$oJV.JournalEntries.Lines.AccountCode='252000-01'
$oJV.JournalEntries.Lines.Credit = 200

$oJV.Add() |Out-Null
$oCompany.GetLastErrorDescription()
