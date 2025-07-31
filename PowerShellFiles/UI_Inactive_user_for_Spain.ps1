<#
    2022.09.22
    UI_DI_Inactive users in site list

#>

cls
$sitelist = "SAPB1-ESP"
$ticktNum ="INC0291371"

$users_tobeInactive =  @(
'albayoa',
'diazsol',
'albaflo',
'vazquluc',
'lozanant',
'calvomar',
'gomezlui',
'gonzafer',
'uretacar',
'berencar',
'rosaljoa',
'diazson',
'sabatpie',
'campopab',
'campajos',
'gutieeli',
'durremer',
'fernaadr',
'gonzaemi',
'gadewa',
'gaskapau',
'espinjua',
'andruolg',
'alvarroc',
'cachafra',
'huertman',
'lastrjos',
'llanejav',
'marcosau',
'mecncprd',
'mepress',
'suarelui',
'lablendlab',
'meplateprd',
'Urbanmar',
'fernalui',
'karapjoa',
'postrelz',
'lipinsla',
'syreklen')


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
		 Function ConnectToCompany {
       Return $oCompany.Connect()
    }
function UI_DI_Conn
{


    # Connect to SBO via DI API

   

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


 Start-Sleep 5                                  # 开始连接DI

        if (SetConnectionContext -ne 0 ) {$SBO_Application.MessageBox("Failed setting a connection to DI API");Exit}
        if (ConnectToCompany -ne 0 ) {$SBO_Application.MessageBox("Failed connecting to the company's Data Base") ; Exit}
        Write-Host -BackgroundColor Cyan "DI Connected To: " $oCompany.CompanyName

 foreach($USR in $users_tobeInactive)                # 循环登录不同的 COMPANY DB
{
  
       
      

          #disable SP contorl in TNMSGS UDO
       # Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"
        # fn_SAPB1_SP_control $ticktNum 'N'

          # ####################################
          #  批量查找用户,移除 Domain account ,加上  Inactive -  并lock
          # ####################################
        
       $sitec =  $site.Substring($site.Length-2)
          
        

           $SBO_Application.ActivateMenuItem("8449")  # Open user set up form

          $SBO_Application.ActivateMenuItem("1281")   # enter Find mode

          $oUserForm =  $SBO_Application.Forms.ActiveForm
  
          $oUserForm.Items.Item(“13”).specific.value =    $USR

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



# remove user access 

	$users_remove_access = @('albaflo',
'albayoa',
'ALONSDON',
'alonspur',
'alvarroc',
'andruolg',
'arbesjav',
'attaribr',
'B1i1',
'berencar',
'bernaale',
'cachafra',
'calvomar',
'campajos',
'campopab',
'cevelire',
'chuahivy',
'cisnecar',
'daneksla',
'delcujor',
'dhadaman',
'diazsol',
'diazson',
'durremer',
'espinjua',
'estraest',
'farbipio',
'fauzinin',
'fernaadr',
'fernalui',
'fernanat',
'gadewa',
'gaskapau',
'gomezlui',
'gonzaemi',
'gonzafer',
'grandalf',
'gutieeli',
'hansuh',
'huertman',
'karapjoa',
'kotkrz',
'kulikpaw',
'L2sap',
'l2sap2',
'lastrjos',
'lipinsla',
'llanejav',
'lozanant',
'maluflui',
'manager',
'marcosau',
'mecncprd',
'meplateprd',
'mepress',
'motykjoa',
'postrelz',
'ramoseme',
'rbanmar',
'rosaljoa',
'rostrste',
'sabatpie',
'sanchher',
'sedanana',
'serednat',
'stenpdir',
'suarelui',
'Support',
'syreklen',
'Urbanmar',
'uretacar',
'vazquluc',
'wo',
'WOCJIPAT',
'wojcipat',
'Workflow',
'yaogav',
'zhukev',
'ziebamon',
'daneksla',
'harazjul',
'janusbar',
'kolodpau',
'michaang',
'pawelmat',
'przewdom',
'Przewirska',
'srokamon',
'srokapio',
'suchamar',
'vandeluc',
'zhuali',
'zielekat',
'liskry')
       #open Authorization form
       $SBO_Application.ActivateMenuItem("3332")
        $oForm = $SBO_Application.Forms.ActiveForm
    
         # click tab of user
      $oForm.Items.Item("234000005").Click(0)
        
       $Omatrix = $oForm.Items.Item("5").Specific
      $oCol = $Omatrix.Columns.Item("0")        
        
       $lMtxRow = $Omatrix.VisualRowCount
        
    For ($k = 1 ;$k -le  $lMtxRow ;$k++ ) {
		   foreach ($user in $users_remove_access)
    {
           # click each user
     
        if( $oCol.Cells.Item($k).Specific.Value -eq $user)
        { $oCol.Cells.Item($k).Click(0)

          # click No access button
           $oForm.Items.Item("4").Click(0)

            # click 'update' button
           $oForm.Items.Item("1").Click(0)
         }
        }
   }
           # Enable SP contorl in TNMSGS UDO
        Write-Host  -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =""
       #  fn_SAPB1_SP_control $ticktNum2 'Y'

       #  break  #quit current loop,jump to next site loop
      
     
    
 

    
  