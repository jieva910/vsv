
    # Connect to SBO via UI API
        Function SetApplication {
              $SboGuiApi = New-Object -comobject "SAPbouiCOM.SboGuiApi"
              $sConnectionString =  "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

              $SboGuiApi.Connect($sConnectionString)
              $SboGuiApi.GetApplication()
          
        }


   $SBO_Application = SetApplication

   $SBO_Application.ActivateMenuItem("COR_COR_Administration")  # 打开 Coresuite license 窗口
   $oForm = $SBO_Application.Forms.ActiveForm   

   $oForm.Items.item("modLFold").click(0)                       # 点击 license manage 页签


  $gridUser =  $oForm.Items.Item("usrGrid").specific            # 左边Grid的对象
  $col =  $gridUser.Columns.Item("Usercode") 

  $gridLicense = $oForm.Items.Item("lGrid").specific            # 右边Grid的对象
  $colUseLic = $gridLicense.Columns.Item("Use licence")
  $colCoresuitName = $gridLicense.Columns.Item("Name")

 
$users = @('kosjoa',
'moundrah',
'strycmar',
'janasmat',
'paszkjak',
'wochAnn',
'baranmar',
'Polusann',
'kuszcmag',
'blechagn',
'poluspaw',
'urbanelz',
'morekmag',
'szewcjol',
'zalusagn',
'wroberen',
'kossajak',
'addarmag',
'michakat',
'milonann',
'stelmkin',
'biegaewa',
'kusiaagn',
'kadzimar',
'palkamal',
'franktet',
'sulikadr',
'lagowjac',
'mermomac',
'wanotmag',
'skowrann',
'lewanewa',
'mojecann',
'zaziawoj',
'strusmal',
'kuchamar',
'bodziprz',
'czapijak',
'bogacart',
'kuczyagn',
'gliuzkyr',
'kukkarag',
'kukkarag',
'Budekdan',
'bednagab',
'zbikodar',
'gubalang',
'labedren',
'dzialpau',
'kubicmar',
'jarzykat',
'mlodzmic',
'boudrmat',
'pietrkat',
'herasolh',
'sharuali',
'laskajus',
'sromedom',
'kuzawoj',
'checolg',
'chorzadr',
'radzkjak',
'rafaalb',
'bajorann',
'bondakar',
'bondaboh',
'romanser',
'rosakat',
'srokamon',
'fryzliza',
'frydrann',
'Zmudapat',
'grzezrob',
'kaczmkat',
'malinann',
'langejak',
'harazjul',
'gedledom',
'glegodia',
'tomczewe',
'bialabar',
'dziewmar',
'jackidaw',
'korczali',
'brozddor',
'pogodluk',
'porebdor',
'hudziann',
'Lippakat',
'sliwailo',
'pawlokon',
'urbanbar',
'cygangab',
'smietjoz',
'jieva',
'fudaledy')

foreach($u in $users)
{
 for( $i= 0 ;$i -lt $gridUser.Rows.Count;$i+=1 ){
     $user =  $col.GetText($i)
    if( $user -match $u) {                                # 找到需要分配license的用户
       $col.Click($i,1,1)                                      # grid 的 column 没有cell对象,点击动作必须使用col对象方法
        break
    }
  }

 for ($j=0;$j -lt  $gridLicense.Rows.Count ; $j+=1){
    
   if ( $colCoresuitName.GetText($j) -match "coresuite designer"){
     
        if (!$colUseLic.IsChecked($j))                             # 忽略已分配license用户
         { $colUseLic.Click($j,0,1) }                              # 中间的 0 表示勾选，1表示去掉勾
        break
   }
     
 }
}



