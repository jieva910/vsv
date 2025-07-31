

<#
   purpose: 自定义选择 quick copy 里面的 BP 主数据
    2022.09.06
#>
    # Connect to SBO via UI API
        Function SetApplication {
              $SboGuiApi = New-Object -comobject "SAPbouiCOM.SboGuiApi"
              $sConnectionString =  "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

              $SboGuiApi.Connect($sConnectionString)
              $SboGuiApi.GetApplication()
          
        }


   $SBO_Application = SetApplication

 #  $SBO_Application.ActivateMenuItem(8208)                        
    $oForm = $SBO_Application.Forms.ActiveForm
 
 # BP master data Matrix inside Quick copy 
    $Mtx = $oForm.Items.Item("1470000015").Specific
                        
   
    $users = @('COGP',
'Country',
'EE',
'ENG',
'FSAS',
'FSASCU',
'FSHS',
'GA',
'Global',
'HQ',
'HR',
'HSE',
'INDE',
'IT',
'ITTEGAP',
'ITTEOP',
'LIGL',
'LISS',
'Local',
'LOTM',
'MEAL',
'MT',
'NOPE',
'OPMA',
'OTCO',
'OTGL',
'PASU',
'PASUCO',
'PASULEAN',
'POA',
'POM',
'PRMAAR',
'PROD',
'PURC',
'RATI',
'RD',
'Regional',
'REVE',
'SEL',
'SHTM',
'XX',
'XXXX'
    )
    

$arrlist =[System.Collections.ArrayList]::new()                # 存入数组，用以统计
  

foreach($u in $users)
{
    for($i=1;$i -lt $Mtx.VisualRowCount ; $i+=1){                # 循环查找
      
      if ($Mtx.Columns.Item("_FU_P_2").Cells.Item($i).Specific.Value -eq $u) {
              $Mtx.Columns.Item("1470000001").Cells.Item($i).Specific.Checked = 1      # 勾选框               
                 [void]$arrlist.Add($u)                               #存入数组，用以统
          break                                                     # 退出循环
      }
    }
}

# $oForm.Items.Item("1").click(0)                                  # license 分配完成点击更新

write-host -background DarkCyan ” $arrlist has been selected"

$arrlist.Clear()                                                 # 释放内存

