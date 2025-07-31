

$oCompany = New-Object -COMobject "SAPbobsCOM.Company"
$SBO_Application = ""
$ticktNum="INC0119233"

# Connect to SBO via UI API
    Function SetApplication {
          $SboGuiApi = New-Object -comobject "SAPbouiCOM.SboGuiApi"
          $sConnectionString =  "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

          $SboGuiApi.Connect($sConnectionString)
          $SboGuiApi.GetApplication()
          
    }

#Connect with connection string

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

# Initialiation SBO 
 $SBO_Application = SetApplication
if (SetConnectionContext -ne 0 ) {$SBO_Application.MessageBox("Failed setting a connection to DI API")}
if (ConnectToCompany -ne 0 ) {$SBO_Application.MessageBox("Failed connecting to the company's Data Base") }
Write-Host "DI Connected To: " $oCompany.CompanyName
#$SBO_Application.MessageBox("DI Connected To: " + $oCompany.CompanyName)


 #disable SAPB1 TN sp control


 #reset current user's password
$usr = $oCompany.GetBusinessObject(12)
 $rs = $oCompany.GetBusinessObject('300') #recordset
 $usercode = $oCompany.UserName
    $rs.doquery(¡°select userid from ousr where user_code='" +$usercode + "'")
    if ($rs.EoF -eq $false)
     {  $uid = $rs.Fields.Item(0).value
        if ($usr.getbykey($uid) -eq $true )
        { #$usr.locked =0 ;
            $usr.UserPassword ='Ves-123456'
          #$usr.SUPERUSER=1
            $usr.Update()
            $oCompany.GetLastErrorDescription()
            }}

#Enable sapb1 TN SP control
 

 $oCompany.Disconnect()