<#
      3个月没有登陆的用户，账号锁定
      如果域账号存在，命名locked - user name 
      如果不存在，命名 Left - user name

#>

cls
$sitelist = "SAPB1_AS","SAPB1_BY","SAPB1_YK","SAPB1_CS","SAPB1_WN","SAPB1_KT","SAPB1-SZ","SAPB1_SQ","SAPB1_WE"
$ticktNum ="inactive user by audit request"

$users_tobeInactive = "SELECT t.USERID,t.USER_CODE,t.U_NAME,isnull(t.lastLogin,'19000101') lastLogin,isnull(t.LstLogoutD,'19000101') LstLogoutD,t.Locked,t.DomainUser,t.createDate FROM OUSR t where t.Locked = 'n' 
  AND t.GROUPS <> 99
  AND (datediff(mm,isnull(t.lastLogin,'20050101'),getdate())>=3 or datediff(mm,isnull(t.LstLogoutD,'20050101'),getdate())>=3 )   -- not log in sap within 3 months
 AND datediff(dd,isnull(t.createDate,'20050101'),getdate()) >= 30  
 AND t.USER_CODE NOT IN ('reddyram','peddipra','jhapan','moundrah','AlertSvc','Halo','TITUser','montova','Workflow','ClosePeriod')
 and t.superuser = 'N'
"


# no AD account users who log in  within 30 days  
$b1UserWithDomainAcct ="SELECT t.user_code,REPLACE(t.DomainUser,'corp\','') DomainUser,t.lastlogin FROM ousr t WHERE isnull(t.DomainUser,'') <>'' and GROUPS <> 99"

    
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

# LDAP Query
Function Fn_Return_AD_Usr_infor ($sUsrName)
{
   $Searcher = New-Object DirectoryServices.DirectorySearcher
        $Searcher.SearchRoot = 'LDAP://DC=corp,DC=vesuvius,DC=com'
        $Searcher.Filter = '(&(objectCategory=user)(cn='+$sUsrName+'))'
        $res = $Searcher.FindAll()  | Sort-Object path
   $Value = "" | Select-Object -Property User_code,first_name ,last_name,email,displayName,phonenum,empid
        
        foreach ($usrTmp in $res)
        {  $Value.User_code = $usrTmp.Properties["name"]
           $Value.first_name = $usrTmp.Properties["sn"] 
            $Value.last_name  = $usrTmp.Properties["givenname"]
            $Value.email = $usrTmp.Properties["mail"] 
            $Value.displayName= $usrTmp.Properties["displayName"]
            $Value.phonenum=$usrTmp.Properties["telephoneNumber"]
            $Value.empid=$usrTmp.Properties["employeeid"]

          #$usrtmp.Properties
        }

  Return $Value
}



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

         $oRs = $oCompany.getbusinessobject(300)
         $oRs.doquery($users_tobeInactive)

         if (!$ors.EoF)
            {
             [xml]$xmls = $ors.GetAsXML()
             $nodes = $xmls.SelectNodes("//row")
             foreach($r in $nodes)
             {

               $SBO_Application.ActivateMenuItem("8449")  # Open user set up form
              $SBO_Application.ActivateMenuItem("1281")   # enter Find mode
              $oUserForm =  $SBO_Application.Forms.ActiveForm 
              $oUserForm.Items.Item(“13”).specific.value =    $r.USER_CODE
              $ouserform.Items.item(“1”).click(0)                     # 点击查找
                  $username = $oUserForm.Items.Item("14").specific.value
                   $v = Fn_Return_AD_Usr_infor $r.USER_CODE
                    if ($v.User_code -eq $null){$oUserForm.Items.Item("14").specific.value = "Left - $($username)” ; $oUserForm.Items.Item(“1470000131”).specific.value = “” }
                   else{ $oUserForm.Items.Item("14").specific.value = "Locked - $($username)” }                
                  $locked = $oUserForm.Items.Item("10000116").specific
                  if(!$locked.Checked){$locked.Item.click(0)}            # 选中 Locked
                   $ouserform.Items.item(“1”).click(0)                      # 点击 更新

                   "$($oUserForm.Items.Item("14").specific.value)  in $($oCompany.CompanyDB)"  +" last login $($r.lastLogin) " | Out-File  $logPath -Append
                   $ouserform.Items.item(“2”).click(0)                      # 点击 取消

            }
          }
           

          $oRs2 = $oCompany.getbusinessobject(300)
         $oRs2.doquery($b1UserWithDomainAcct)
         if (!$oRs2.EoF)
            {
             [xml]$xmls2 = $oRs2.GetAsXML()
             $nodes2 = $xmls2.SelectNodes("//row")
             $users_noAD = foreach($r2 in $nodes2)
              {
               $v2 = Fn_Return_AD_Usr_infor $r2.DomainUser
               if ($v2.User_code -eq $null)
               { "'"+$r2.USER_CODE+"'" }
              }
             }
        if($users_noAD)
		{
        $sql_noADusers = "select user_code,isnull(lastLogin,'19000101') lastLogin from ousr where  user_code in ($($users_noAD -join ","))"
        $oRs3 =  $oCompany.getbusinessobject(300)
         $oRs3.doquery($sql_noADusers)
         if (!$oRs3.EoF)
             { 
                [xml]$xmls3 = $oRs3.GetAsXML()
                 $nodes3 = $xmls3.SelectNodes("//row")
               foreach($r3 in $nodes3)
              {
               $SBO_Application.ActivateMenuItem("8449")  # Open user set up form
              $SBO_Application.ActivateMenuItem("1281")   # enter Find mode
              $oUserForm =  $SBO_Application.Forms.ActiveForm
              $oUserForm.Items.Item(“13”).specific.value = $r3.USER_CODE
              $ouserform.Items.item(“1”).click(0)                     # 点击查找
                  $username = $oUserForm.Items.Item("14").specific.value
                  $oUserForm.Items.Item("14").specific.value = "Left - $($username)” ; $oUserForm.Items.Item(“1470000131”).specific.value = “” 
                  $locked = $oUserForm.Items.Item("10000116").specific
                  if(!$locked.Checked){$locked.Item.click(0)}            # 选中 Locked
                   $ouserform.Items.item(“1”).click(0)                      # 点击 更新
                  "$($oUserForm.Items.Item("14").specific.value)  in $($oCompany.CompanyDB)"  +" last login $($r3.lastLogin) " | Out-File  $logPath  -Append
                   $ouserform.Items.item(“2”).click(0)                      # 点击 取消
            }
          }
		  }
           

        $ticktNum2 =""
         fn_SAPB1_SP_control $ticktNum2 'Y'  $site

         break  #quit current loop,jump to next site loop
      
     }
    }
 }
 
   $oCompany.Disconnect()
 if($logPath){ Send-MailMessage -Body 'check attachment' -From AUDIT@SAPB1APP.COM  -Subject 'sapb1 users locked' -To 'evan.ji@vesuvius.com' -Attachments $logPath  -SmtpServer 'APMailrelay.vesuvius.com' -port 25  }
taskkill /f /t /im   'SAP Business One.exe'
 Stop-Process -Id $PID
 
