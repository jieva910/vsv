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

# update Valid to  special price for SAPB1_OS

$SQL = "select t.ItemCode,t.CardCode from spp1 t where t.ItemCode in 
(select t.ItemCode from oitw t where t.WhsCode = 'os-o')"

$oRs = $cmp.GetBusinessObject(300)
 $oSpecialPrices = $cmp.GetBusinessObject(7)    # oSpecialPrices = 7

$oRs.DoQuery($SQL)

if (!$oRs.EoF){

 [xml]$Data = $oRs.GetAsXML()
 
 $Nodes = $Data.SelectNodes("//row")
 
 foreach($node in $Nodes)
 {  
   
    if ($oSpecialPrices.GetByKey($node.ItemCode,$Node.CardCode))
         {
            $lastline  = $oSpecialPrices.SpecialPricesDataAreas.Count - 1 
            $oSpecialPrices.SpecialPricesDataAreas.SetCurrentLine($lastline)
            $oSpecialPrices.SpecialPricesDataAreas.Dateto ='2025-01-01'
         Write-Host  $node.ItemCode,$Node.CardCode  $oSpecialPrices.Update() $cmp.GetLastErrorDescription()
          
         }
   
    
 } 

}





  
