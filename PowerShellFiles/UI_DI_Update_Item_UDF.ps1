

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
    Write-Host -ForegroundColor Cyan "DI Connected To: " $cmp.LicenseServer $cmp.CompanyName
}

UI_DI_Conn



 $oitemcode=$cmp.GetBusinessObject(4)

 $csv_csitem =  import-csv \\dg-fs-usr02\KKUsers\jieva\temp\CLCommodityCorrections.csv


 $starttime = Get-Date 

  foreach($r in $csv_csitem)
 { 

   if ($oitemcode.GetByKey($r.'ItemCode')){
   
      if ( $oitemcode.UserFields.Fields.Item("U_Ves_PRP1").value -ne  $r.U_Ves_PRP1 -and $oitemcode.UserFields.Fields.Item("U_Ves_PRP2").value -ne $r.U_Ves_PRP2)
       {
       $oitemcode.UserFields.Fields.Item("U_Ves_PRP1").value =$r.U_Ves_PRP1
       $oitemcode.UserFields.Fields.Item("U_Ves_PRP2").value =$r.U_Ves_PRP2
       }
    }
  $cmp.CompanyDB + $r.'ItemCode' +' updated  with exception:0 with error code:'+ $oitemcode.Update() + ' and error description is:'+  $cmp.GetLastErrorDescription() 


 }

 $endtime=  Get-Date

 "running time:" + ($endtime-$starttime).Seconds

 
