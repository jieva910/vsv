# for upgrade to SQL SERVER 2016
cls
$sitelist =  "SAPB1_SZ","SAPB1_CS","SAPB1_WN","SAPB1_KT","SAPB1_AS","SAPB1_BY","SAPB1_YK","SAPB1_HG"  #qianbry2

$csv = Import-Csv 'C:\Dell\all china SAP USERS with domain account.csv'

$ticktNum=  "SAP db upgrade to sql2016" 

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




function Fn_Assign_User_DomainAccount ($sUsrName,$domainaccount)
{
  
    $SBO_Application.ActivateMenuItem("8449") # open User setup form
    $SBO_Application.ActivateMenuItem("1281") # open Find user mode
    $oForm = $SBO_Application.Forms.ActiveForm
     $oItem = $oForm.Items.Item("13")
    $oItemAD = $oForm.Items.Item("1470000131") # field for user's bind with windows account
    $oItem3 = $oForm.Items.Item("231000004") # field for Authorisation Groups

    $oItem.Specific.Value = $sUsrName
    $oForm.Items.Item("1").Click(0)

    $oItemAD.Specific.Value = "$domainaccount"
    
    #update user
    $oform.Items.Item("1").Click(0)
  
}






# ############################### Initialiation SBO 
 #$SBO_Application = SetApplication



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

        #start assign domain account to  user..
        foreach($r in $csv | ?{$_.sitecode -match $site.Substring($site.Length-2,2)})
        {
        Fn_Assign_User_DomainAccount $r.USER_CODE $r.domainuser
       }
        Start-Sleep  5 
         # Enable SP contorl in TNMSGS UDO
        Write-Host  -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =""
         fn_SAPB1_SP_control $ticktNum2 'Y'

         break  #quit current loop,jump to next site loop
       }
      
    
   }
}
