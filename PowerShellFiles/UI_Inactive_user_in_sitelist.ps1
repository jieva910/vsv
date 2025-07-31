
. C:\shared\PShell\PSLib\SAPB1_UI_DI_Connection.ps1

$sitelist = "SAPB1_AS","SAPB1_BY"
$ticktNum ="INC0201348"
 $SBO_Application = SetApplication

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

          #disable SP contorl in TNMSGS UDO
        Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"
         fn_SAPB1_SP_control $ticktNum 'N'

          # ####################################
          #  批量查找用户并lock
          # ####################################
           
         $csv = Import-Csv    C:\Temp\user.csv

         $sitec =  $site.Substring($site.Length-2)
        
         foreach($r in $csv|?{ $_.'Site code' -EQ $sitec})
         { 
  
           $SBO_Application.ActivateMenuItem("8449")  # Open user set up form

          $SBO_Application.ActivateMenuItem("1281")   # enter Find mode

          $oUserForm =  $SBO_Application.Forms.ActiveForm
  
          $oUserForm.Items.Item(“13”).specific.value =  $r.USER_CODE

          $ouserform.Items.item(“1”).click(0)                     # 点击查找
  
           $locked = $oUserForm.Items.Item("10000116").specific
               if(!$locked.Checked) {  # IF USER not locked
              $username = $oUserForm.Items.Item("14").specific.value

              $oUserForm.Items.Item("14").specific.value = "Left - $($username)”

              $oUserForm.Items.Item(“1470000131”).specific.value = “”  # 用户绑定windows域账号

              
            $locked.Item.click(0)            # 选中 Locked

               $ouserform.Items.item(“1”).click(0)                      # 点击 更新

               Write-Host "User Code :$($oUserForm.Items.Item("13").specific.value) updated in $($oCompany.CompanyDB) with err description: "$oCompany.GetLastErrorDescription()
               $ouserform.Items.item(“2”).click(0)                      # 点击 取消
               }

          }



           # Enable SP contorl in TNMSGS UDO
        Write-Host  -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =""
         fn_SAPB1_SP_control $ticktNum2 'Y'

         break  #quit current loop,jump to next site loop
      
     }
    }
 }

    
  
       
    