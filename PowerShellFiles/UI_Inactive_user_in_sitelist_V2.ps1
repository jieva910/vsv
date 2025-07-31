<#
    2022.09.22
    UI_DI_Inactive users in site list

#>

cls
$sitelist = "SAPB1_BY","SAPB1_WN","SAPB1_SQ"
$ticktNum ="inactive user by julia request"

$users_tobeInactive = import-csv "\\Client\C$\Temp\chinaInactive Users Request.csv"

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

UI_DI_Conn


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
        Write-Host -BackgroundColor Cyan "DI Connected To: " $oCompany.CompanyName

          # ####################################
          #  批量查找用户,移除 Domain account ,加上  Inactive -  并lock
          # ####################################
        
       $sitec =  $site.Substring($site.Length-2)
                 #disable SP contorl in TNMSGS UDO
        Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"
         fn_SAPB1_SP_control $ticktNum 'N'  $sitec
          
        foreach($r in $users_tobeInactive|?{ $_.'Site code' -EQ $sitec}){

           $SBO_Application.ActivateMenuItem("8449")  # Open user set up form

          $SBO_Application.ActivateMenuItem("1281")   # enter Find mode

          $oUserForm =  $SBO_Application.Forms.ActiveForm
  
          $oUserForm.Items.Item(“13”).specific.value =    $r.USER_CODE

          $ouserform.Items.item(“1”).click(0)                     # 点击查找
  

              $username = $oUserForm.Items.Item("14").specific.value

              $oUserForm.Items.Item("14").specific.value = "Left - $($username)”

              $oUserForm.Items.Item(“1470000131”).specific.value = “”  # 用户绑定windows域账号

              $locked = $oUserForm.Items.Item("10000116").specific
    
              if(!$locked.Checked){$locked.Item.click(0)}            # 选中 Locked

               $ouserform.Items.item(“1”).click(0)                      # 点击 更新

               Write-Host "User Code :$($oUserForm.Items.Item("13").specific.value) updated in $($oCompany.CompanyDB) with err description: "$oCompany.GetLastErrorDescription()
               $ouserform.Items.item(“2”).click(0)                      # 点击 取消


          }


           # Enable SP contorl in TNMSGS UDO
        Write-Host  -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =""
         fn_SAPB1_SP_control $ticktNum2 'Y'  $sitec

         break  #quit current loop,jump to next site loop
      
     }
    }
 }

    
  