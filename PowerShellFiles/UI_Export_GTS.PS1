

cls
$sitelist = "SAPB1_AS","SAPB1_BY","SAPB1_YK","SAPB1_CS","SAPB1_WN","SAPB1_KT","SAPB1-SZ"
$ticktNum ="auto Export sales invoices for GTS(Shu Dian fa piao) "


    
 $logPath = "d:\audit_log\user_tobelocked_log_"+(Get-Date -Format 'yyyyMMddhhmmms').ToString()+".txt"

$oCompany = new-object -ComObject "SAPBOBSCOM.Company"

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
   
        if (SetConnectionContext -ne 0 ) {$SBO_Application.MessageBox("Failed setting a connection to DI API");break}
        if (ConnectToCompany -ne 0 ) {$SBO_Application.MessageBox("Failed connecting to the company's Data Base") ; break}
    Write-Host -ForegroundColor Cyan "DI Connected To: " $oCompany.LicenseServer $oCompany.CompanyName
}


$loginForm = $SBO_Application.Forms.ActiveForm

$loginForm.Items.item(“1”).click(0)  

start-sleep 5


# Control the SP of blocking Super user
Function fn_SAPB1_SP_control($ticknum,$YesNO,$site)
  {   
   $oCompServic = $oCompany.GetCompanyService()
   $oGeneralServic = $oCompServic.GetGeneralService('VES_TNMSGS')
   $oGeneralParams = $oGeneralServic.GetDataInterface(3)  # gsGeneralDataParams = 3
   $oGeneralParams.SetProperty('Code','9999999')
   $oGeneralData = $oGeneralServic.GetByParams($oGeneralParams)
   
   # check current SP control status
   $isActive =  $oGeneralData.GetProperty('U_VES_Active')
   $comment = $ticknum
   if ($isActive -ne $YesNO)
     {
          switch ($site)
           {  { $site -in "xx"} {$ogeneraldata.SetProperty('U_VES_COmments',$comment)} 
     
              Default { $ogeneraldata.SetProperty('U_VES_Comments',$comment)}
           }
           $ogeneraldata.SetProperty('U_VES_Active',$YesNO)
           $oGeneralServic.Update($ogeneraldata)

           Release-Ref ($oCompServic)
      }
}


# Release COM object 
function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}




 foreach($site in $sitelist)                # 循环登录不同的 COMPANY DB
{
  $SBO_Application.ActivateMenuItem("3329") #open choose company list

  $oForm = $SBO_Application.Forms.ActiveForm
    $oItem = $oForm.Items.Item("4")
    $Omatrix = $oItem.Specific
    $oChkbox = $oForm.Items.Item("1470000128").Specific
    $strDBSvr = $oForm.Items.Item("420000125").Specific.Value
   Foreach ( $j in  1..$Omatrix.VisualRowCount)
    { 
      $strCompname = $Omatrix.Columns.Item("2").Cells($j).Specific.Value
    If ($strCompname -eq $site  ) {
       If (!$oChkbox.Checked) { $oChkbox.Item.Click(0)}
       $Omatrix.Columns.Item("2").Cells($j).Click(1)  #log on the company db with windows domain account
      
       #break                                          # 成功登录之后，退出当前循环
       
       Start-Sleep 5                                  # 开始连接DI

        if (SetConnectionContext -ne 0 ) {$SBO_Application.MessageBox("Failed setting a connection to DI API");Exit}
        if (ConnectToCompany -ne 0 ) {$SBO_Application.MessageBox("Failed connecting to the company's Data Base") ; Exit}
      

        Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"
         fn_SAPB1_SP_control $ticktNum 'N'  $site

        ## STRAT TASK
		   
  Add-Type @"
using System;
using System.Runtime.InteropServices;
public class SFW {
 [DllImport("user32.dll")]
 [return: MarshalAs(UnmanagedType.Bool)]
 public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@

$fw =  (get-process 'SAP Business ONE').MainWindowHandle
[SFW]::SetForegroundWindow($fw)

Start-Job -ScriptBlock {$rtTrue =$false
do 
{
    $filename = Get-Date -Format 'yyyyMMddhhmmssm'
$wshell = new-object -com wscript.shell
$rtTrue = $wshell.appactivate("Save As")
if($rtTrue){$wshell.sendkeys("$($filename)");$wshell.sendkeys("{Enter}");break}
   Start-Sleep -s 10
} until  ($rtTrue -eq $ture)
}

  # open Export to GTS form
       $SBO_Application.ActivateMenuItem("13569")
        $oForm = $SBO_Application.Forms.ActiveForm
        
        #  posting date from
        $oForm.Items.Item("1230000010").Specific.Value ='20230612'  # GET-DATE -FORMAT 'yyyyMMdd'  #  
        
        #  posting date to
        $oForm.Items.Item("1230000012").Specific.Value ='20230812' # GET-DATE -FORMAT 'yyyyMMdd' # 
        
       
       #  click ok button
        $oForm.Items.Item("1230000001").Click(0)
        
        #  document to be exported form
        $oForm_doc = $SBO_Application.Forms.ActiveForm
        
        $oChkbox = $oForm_doc.Items.Item("1230000004").Specific
         If (!$oChkbox.Checked) { $oChkbox.Item.Click(0)  }
        $oMaxtirx = $oForm_doc.Items.Item("1230000003").Specific

          $lMtxRow = $oMaxtirx.VisualRowCount
       
          For( $k = 1 ;$k -le $lMtxRow;$k++)
          {
                #  selected checkbox
                $oMaxtirx.Columns.Item("_FU_P_2").Cells($k).Click(0)
                
               #  oWs.Cells(k + 4, "B") = $oMaxtirx.Columns.Item("OINV_1").Cells(CInt(k)).Specific.Value
                #  selected document NO. for next query
              $strDocnum +=  $oMaxtirx.Columns.Item("OINV_1").Cells($k).Specific.Value 
         
          }
 
        if ($strDocnum)
         {
          $oForm_doc.Items.Item("1230000001").Click(0)
         
         }
        else {$oForm_doc.Items.Item("1230000002").Click(0)}
                 
         
		## END TASK
           

        $ticktNum2 =""
         fn_SAPB1_SP_control $ticktNum2 'Y'  $site

         break  #quit current loop,jump to next site loop
      
     }
    }
 }
 
   $oCompany.Disconnect()
 if($logPath){ Send-MailMessage -Body 'check attachment' -From ExportGTS@SAPB1APP.COM  -Subject 'Auto GTS Export' -To 'evan.ji@vesuvius.com' -Attachments $logPath  -SmtpServer 'APMailrelay.vesuvius.com' -port 25  }
taskkill /f /t /im   'SAP Business One.exe'
 Stop-Process -Id $PID
 
