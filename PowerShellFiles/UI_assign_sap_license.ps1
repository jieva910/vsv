
    # Connect to SBO via UI API
        Function SetApplication {
              $SboGuiApi = New-Object -comobject "SAPbouiCOM.SboGuiApi"
              $sConnectionString =  "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

              $SboGuiApi.Connect($sConnectionString)
              $SboGuiApi.GetApplication()
          
        }


   $SBO_Application = SetApplication

    $SBO_Application.ActivateMenuItem(8208)                        # open B1 License form
    $oForm = $SBO_Application.Forms.ActiveForm

    $ItemUser = $oForm.Items.Item("4").specific                    # allocation 页签下面的 User item
    $oColuser = $ItemUser.Columns.Item("1")                        # User column

    $ItemB1lic = $oForm.Items.Item("18").specific                  # allocation 页签下面的 B1 user type license
    $oColB1lic = $ItemB1lic.Columns.Item("10")                     # 1st column
    $oColB1licused = $ItemB1lic.Columns.Item("6")                  # 2nd column
  
    $users = @('yuanwen','zhangcan','zhouxia'                      # SAPB1 用户列表
    ,'zhangsum','zhangjoy','B1i','buttnas'
    )
    

    $arrlist =[System.Collections.ArrayList]::new()                # 将有licens的 用户存入数组，用以统计

    

foreach($u in $users)
{
    for($i=1;$i -lt $oColuser.Cells.Count ; $i+=1){                # 循环查找用户
      
      if ($oColuser.Cells.Item($i).specific.value -match $u) {
          $oColuser.Cells.Item($i).click(0)                             # 选中用户

           for($j = 1;$j -lt $oColB1lic.Cells.Count;$j+=1)
           {
             if ($oColB1lic.Cells.Item($j).specific.value -match "SAP Business One Professional User")
             {
               $oCheck =  $oColB1licused.Cells.Item($j).specific       # 勾选框   
                
              if (!$ocheck.Checked) {
                $oCheck.Checked = $true                               # 忽略已分配License 用户
                 [void]$arrlist.Add($u)                               # 将有licens的 用户存入数组，用以统计
              }          
              
             }
           }
          break                                                     # 退出循环
      }
     
    }
}

$oForm.Items.Item("1").click(0)                                  # license 分配完成点击更新

write-host -background DarkCyan ”User $arrlist has been assigned license"

$arrlist.Clear()                                                 # 释放内存

    