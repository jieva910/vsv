#**********************
# purpose : add or update User and it's authoriation group via UI,DI API
# Date    : 2020.10
#**********************
cls
$sitelist =  "SAPB1_WN","SAPB1_CS","SAPB1_KT","SAPB1_AS","SAPB1_BY","SAPB1_YK","SAPB1-SZ"  #qianbry
$Continue_withNonAD ="n"

$sUsrName=read-host  "Please Enter New UserCode(5+3)"   # 1 input new usercode 5+3 same as AD account
$sReferencUsrCode = Read-Host "Please enter Reference User code"    # 2) input reference to whom 
$ticktNum= Read-Host "Please Enter ticket number" 

$sSql_userdefault ="SELECT t.DfltsGroup FROM OUSR t WHERE t.USER_CODE = '$sReferencUsrCode'"
$sSql_UserGroup ="SELECT distinct g.GroupName FROM ousr t LEFT JOIN USR7 u ON t.USERID = u.UserId INNER JOIN ougr g ON u.GroupId = g.GroupId WHERE t.USER_CODE = '$sReferencUsrCode'"


    $oCompany = New-Object -COMobject "SAPbobsCOM.Company"
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

          # Release-Ref ($oCompServic)
      }
}

function get_ReferenceUserProperty ([string]$sql)
{
  $oRs = $oCompany.GetBusinessObject('300') #recordset

   $oRs.doquery($sql)

  $queryValue = while(!$oRs.EoF)
  { $oRs.Fields.Item(0).Value;$oRs.MoveNext()}

  return $queryValue
}


function fn_AddUser ($sUsrName)  {

      $oUsr=$oCompany.getbusinessobject(12)

      $Val = Fn_Return_AD_Usr_infor $sUsrName 
     $UserDefault = get_ReferenceUserProperty $sSql_userdefault

     if ($Continue_withNonAD -EQ "Y" ) {
     
        $oUsr.UserCode = $sUsrName 
        $oUsr.UserName = $sUsrName 
        $oUsr.Defaults = $UserDefault 
	    $oUsr.UserPassword = "Ves-123456"
        $intRetCode = $oUsr.Add()
        $outlog =$oCompany.CompanyDB + ' '+ $sUsrName + ' Added User  with error code:' +  $intRetCode + ' and error description is:' + $oCompany.GetLastErrorDescription()
      
     }

     if ($val.User_code -ne $null ) {
         
         $oUsr.UserCode = $Val.User_code[0]
	     $oUsr.UserName = $Val.displayName[0]
	     $oUsr.Email = $Val.email[0]
         $oUsr.Defaults = $UserDefault 
	     $oUsr.UserPassword = "Ves-123456"
	     $intRetCode = $oUsr.Add()
 	     if ($intRetCode -eq 0 )
	     { 
		    $oEmployee = $oCompany.GetBusinessObject(171)
		    $oEmployee.FirstName =  $Val.first_name[0]
		    $oEmployee.LastName = $Val.last_name[0]
		    $oEmployee.ExternalEmployeeNumber =  $Val.empid[0]
		    $oEmployee.ApplicationUserID =$oCompany.GetNewObjectKey()
		    $oEmployee.OfficePhone = $Val.phonenum[0]
		    $oEmployee.Email = $Val.email[0]
		    $oEmployee.Active = 1
		    $outputcode=$oEmployee.Add()
            $outlog =$oCompany.CompanyDB + ' '+ $Val.User_code[0] + ' Added Employee with error code:' +   $outputcode + ' and error description is:' + $oCompany.GetLastErrorDescription()
           
           # assign Authorization grps to new user
           Grant_User_AuthorizationGrp $sUsrName
          }
	      Else {$outlog =$oCompany.CompanyDB + ' '+ $Val.User_code[0] + ' Added User  with error code:' +  $intRetCode + ' and error description is:' + $oCompany.GetLastErrorDescription()}
      }
      
       return $outlog
}

function Grant_User_AuthorizationGrp ($sUsrName)
{
   # if use non AD then no need AGrps
   if ($Continue_withNonAD -ne  "Y" )
   { $SBO_Application.ActivateMenuItem("8449") # open User setup form
    $SBO_Application.ActivateMenuItem("1281") # open Find user mode
    $oForm = $SBO_Application.Forms.ActiveForm
     $oItem = $oForm.Items.Item("13")
    $oItemAD = $oForm.Items.Item("1470000131") # field for user's bind with windows account
    $oItem3 = $oForm.Items.Item("231000004") # field for Authorisation Groups

    $oItem.Specific.Value = $sUsrName
    $oForm.Items.Item("1").Click(0)

    # open Authorization grp
    $oForm.Items.Item("231000004").Click(0)
    $oForm2 = $SBO_Application.Forms.ActiveForm
    $oItem2 = $oForm2.Items.Item("231000003")
    $Omatrix2 = $oItem2.Specific
    
    # retrieve Authroziation groups of Referencing User
    $AuthorizGrps = get_ReferenceUserProperty $sSql_UserGroup

    # loop Authorization gprs and assign to new user
    foreach ($ag in $AuthorizGrps )
    {
      foreach( $j in  1..$Omatrix2.VisualRowCount)
      {
        if ($ag -eq $Omatrix2.Columns.Item("231000002").Cells($j).Specific.Value)
        { # tick authorizaiton grp
          $Omatrix2.Columns.Item("231000001").Cells($j).click(0)

          break 
         }
      }
    }
    # update user's authorization group after assigning
    $oform2.Items.Item("231000001").Click(0)
    $oform2.Items.Item("231000001").Click(0)
    $oItemAD.Specific.Value = "corp\$sUsrName"
    
    #update user
    $oform.Items.Item("1").Click(0)
  }
}



# ################################ import UI DI connection
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
   $Val = Fn_Return_AD_Usr_infor $sUsrName 
   
   # check new user code if  in the AD 
   if($Val.User_code -eq $null)
   {
    Write-host -ForegroundColor Red  "$sUsrName  can not find in Active Directory"
    $prompt = Read-Host -Prompt " Press any key to continue or Ctrl + c to Stop"
    if([bool]$prompt) {Write-Host 'Continue to add new user code ... ';$Continue_withNonAD = "Y" }
    else {break}
    }



# ############################### Initialiation SBO 
 $SBO_Application = SetApplication



 foreach($site in $sitelist)
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
       $Omatrix.Columns.Item("2").Cells($j).Click(1)  #log on the company db

       Start-Sleep 5

        if (SetConnectionContext -ne 0 ) {$SBO_Application.MessageBox("Failed setting a connection to DI API");Exit}
        if (ConnectToCompany -ne 0 ) {$SBO_Application.MessageBox("Failed connecting to the company's Data Base") ; Exit}
        Write-Host "DI Connected To: " $oCompany.CompanyName

        #disable SP contorl in TNMSGS UDO
        Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"
         fn_SAPB1_SP_control $ticktNum 'N'

        #start Add user..
        fn_AddUser $sUsrName
       
        Start-Sleep  5 
         # Enable SP contorl in TNMSGS UDO
        Write-Host  -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =""
         fn_SAPB1_SP_control $ticktNum2 'Y'

         break  #quit current loop,jump to next site loop
       }
      
    
   }
}
