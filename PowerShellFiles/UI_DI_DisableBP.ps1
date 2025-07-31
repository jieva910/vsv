 
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

# Open UDO TransactionNOtifyMessages Form
# $oForm =  $SBO_Application.OpenForm(0,'VES_TNMSGS','')
  
#  SP of blocking Super user
Function fn_Disable_SPcontrol($ticknum,$YesNO)
  {   
   $oCompServic = $cmp.GetCompanyService()
   $oGeneralServic = $oCompServic.GetGeneralService('VES_TNMSGS')
   $oGeneralParams = $oGeneralServic.GetDataInterface(3)  # gsGeneralDataParams = 3
   $oGeneralParams.SetProperty('Code','9999999')
   $oGeneralData = $oGeneralServic.GetByParams($oGeneralParams)
   
   # check current SP control status
   $isActive =  $oGeneralData.GetProperty('U_VES_Active')
   $comment = $ticknum
   if ($isActive -ne $YesNO)
     {     
        $ogeneraldata.SetProperty('U_VES_Comments',$comment)
        $ogeneraldata.SetProperty('U_VES_Active',$YesNO)
        $oGeneralServic.Update($ogeneraldata)
      }
}  
  
  fn_Disable_SPcontrol $ticknum 'N'
 
  $bp=$cmp.GetBusinessObject(2)
  $csv = Import-Csv \\dg-fs-usr02\KKUSERS\jieva\temp\lyBP.csv
  $errcount = 0
  $correctcount = 0
  
  write-host  -ForegroundColor green Processing $csv.count rows......
  FOREACH($R IN $csv){
  
     IF ($BP.GetByKey($R.SupplierCode))
     {
      IF (!$BP.Frozen) {
      $BP.Valid=0
      $BP.Frozen =1 
      $bp.FrozenRemarks = $ticknum

      $Rtcode = $BP.Update()
      # only display  error records
      if ($Rtcode){$errcount++ ; Write-Host "BP $($R.SupplierCode) UPDATE WITH ERR: $Rtcode $($CMP.GetLastErrorDescription())" }
      else {$correctcount++}
     
       }
     }    
     }
     write-host   "$correctcount updated successfully , $errcount failed "
	 
	 $ticknum = ''
	 fn_Disable_SPcontrol $ticknum 'Y'