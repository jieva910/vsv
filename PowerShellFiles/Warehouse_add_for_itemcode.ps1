 
$ticknum = 'INC0213304'
 
$cmp = New-Object -COMobject "SAPbobsCOM.Company"
function UI_DI_Conn
{
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
             $sCookie = $cmp.GetContextCookie()
             $sConnectionContext = $SBO_Application.Company.GetConnectionContext($sCookie)
             If ($cmp.Connected ){$cmp.Disconnect()}
             return $cmp.SetSboLoginContext($sConnectionContext)
        }

    # Connect to SBO via DI API
    Function ConnectTcmp {
       Return $cmp.Connect()
    }

    # connect to DI 
        if (SetConnectionContext -ne 0 ) {$SBO_Application.MessageBox("Failed setting a connection to DI API");break}
        if (ConnectTcmp -ne 0 ) {$SBO_Application.MessageBox("Failed connecting to the company's Data Base") ; break}
    Write-Host -ForegroundColor Cyan "Connected to LicenseServer: $($cmp.LicenseServer), DB: $($cmp.CompanyDB)"
}

UI_DI_Conn   # connection to DI via UI 



$arr = @('0004921',
'0005025',
'0005906'
)

$oitm = $cmp.GetBusinessObject(4)
foreach($a in $arr)
{if ($oitm.GetByKey($a))
   { # $oitm.WhsInfo.SetCurrentLine($i)      
            $oitm.WhsInfo.add() 
            $oitm.WhsInfo.WarehouseCode ='SZ-B-SR'
         Write-Host  " specific  of " $A " has been added " $oitm.Update() $cmp.GetLastErrorDescription()
               }
}
