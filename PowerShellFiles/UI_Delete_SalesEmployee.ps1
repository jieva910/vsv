

$oCompany = New-Object -COMobject "SAPbobsCOM.Company"

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
    Write-Host -ForegroundColor Cyan "DI Connected To:  $($oCompany.LicenseServer) + $($oCompany.companydb)"
}

UI_DI_Conn



$oSp = $oCompany.GetBusinessObject(53)   # Sales employee object
$ORS = $oCompany.GetBusinessObject(300)  # record set
 
 $ORS.DoQuery(“SELECT T0.[SlpCode] FROM OSLP T0 WHERE T0.[SlpName] like 'eva%'”)

 $OSP.Browser.Recordset = $ORS

$oSp.Browser.MoveFirst()

while (!$oSp.Browser.EoF)
{
  
  if ( $oSp.GetByKey("$($oSp.SalesEmployeeCode)")) {
  
    Write-Host "$($osp.SalesEmployeeName) removed with errcode $($oSP.Remove()) $($ocompany.GetLastErrorDescription())"
    }
  
  $osp.Browser.MoveNext()
}
