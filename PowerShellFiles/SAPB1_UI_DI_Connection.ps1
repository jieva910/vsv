
$oCompany = New-Object -COMobject "SAPbobsCOM.Company"



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
         $sCookie=""
         $sConnectionContext=""
         $lRetCode=0

         $sCookie = $oCompany.GetContextCookie()
         $sConnectionContext = $SBO_Application.Company.GetConnectionContext($sCookie)

         If ($oCompany.Connected ){$oCompany.Disconnect()}
       
         return $oCompany.SetSboLoginContext($sConnectionContext)

    }

# Connect to SBO via DI API

Function ConnectToCompany {

   Return $oCompany.Connect()
}

# Control the SP of blocking Super user
Function fn_SAPB1_SP_control($ticknum,$YesNO)
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
           $ogeneraldata.SetProperty('U_VES_Comments',$comment)
           $ogeneraldata.SetProperty('U_VES_Active',$YesNO)
           $oGeneralServic.Update($ogeneraldata)

           Release-Ref ($oCompServic)
      }
}



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

