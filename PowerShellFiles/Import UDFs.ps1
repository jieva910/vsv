


$cmp = New-Object -comobject  'sapbobscom.company'
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp 'YK'
 $cmp.XmlExportType =3
 # export UDT 

  $oUT = $cmp.GetBusinessObject(153) #Const oUserTables = 153 (&H99)
 
 if ( $out.GetByKey('VES_AUTOGRN')) { $oUT.SaveToFile('c:\temp\VES_AUTOGRN.xml')  }

# export UDF
 $ors = $cmp.GetBusinessObject(300)
 $ors.doquery("select  t.TableID,t.FieldID  from CUFD t where t.TableID = '@VES_AUTOGRN' order by t.FieldID ")

 $xml=[xml]$ors.GetAsXML()

 $nodes = $xml.SelectNodes('//row')

 foreach($n in $nodes)
 {
    $oUF = $cmp.GetBusinessObject(152)
   If ($oUF.GetByKey($n.TableID,$n.FieldID) ){$oUF.SaveXML("c:\temp\$($n.TableID)$($n.FieldID).xml")}
 }
 



 # Import to sapb1 db  ---- CAPEX PAYMENT REQUST

 $sites =@('cstst','kttst')                    # @('SZ','KT','WN','CS','AS','BY','YK')

 FOREACH($site in $sites)
 {
 $cmp = New-Object -comobject  'sapbobscom.company'
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp $site
 $oUF2 = $cmp.GetBusinessObjectFromXML("c:\temp\VES_Payment_phase71.xml",0) #Const oUserFields = 152 (&H98)
write-host  $oUF2.add()  $cmp.GetLastErrorDescription()

$oUF3 = $cmp.GetBusinessObjectFromXML("c:\temp\VES_Payment_phase72.xml",0) #Const oUserFields = 152 (&H98)
write-host  $oUF3.add()  $cmp.GetLastErrorDescription()
  }



  # Import UDF

 $sites =@('cstst','kttst')                    # @('SZ','KT','WN','CS','AS','BY','YK')

 FOREACH($site in $sites)
 {
 $cmp = New-Object -comobject  'sapbobscom.company'
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp $site
 $oUF2 = $cmp.GetBusinessObjectFromXML("c:\temp\VES_Payment_phase71.xml",0) #Const oUserFields = 152 (&H98)
write-host  $oUF2.add()  $cmp.GetLastErrorDescription()

$oUF3 = $cmp.GetBusinessObjectFromXML("c:\temp\VES_Payment_phase72.xml",0) #Const oUserFields = 152 (&H98)
write-host  $oUF3.add()  $cmp.GetLastErrorDescription()
  }


   # Import udt to sapb1 db  ---- AUTO GRN Table in SAPB1_YK
   $oUDT = $cmp.GetBusinessObjectFromXML("c:\temp\VES_AUTOGRN.xml",0) #Const oUserFields = 152 (&H98)
write-host  $oUDT.add()  $cmp.GetLastErrorDescription()

 # after UDT ADDED ,NEED TO RE CONNECT SAPB1 COM OR RELEASE COM


 $oUF2 = $cmp.GetBusinessObjectFromXML("C:\Temp\@VES_AUTOGRN0.xml",0) #Const oUserFields = 152 (&H98)
write-host  $oUF2.add()  $cmp.GetLastErrorDescription()

$oUF3 = $cmp.GetBusinessObjectFromXML("C:\Temp\@VES_AUTOGRN3.xml",0) #Const oUserFields = 152 (&H98)
write-host  $oUF3.add()  $cmp.GetLastErrorDescription()

$oUF4 = $cmp.GetBusinessObjectFromXML("C:\Temp\@VES_AUTOGRN4.xml",0) #Const oUserFields = 152 (&H98)
write-host  $oUF4.add()  $cmp.GetLastErrorDescription()


$oUF5 = $cmp.GetBusinessObjectFromXML("C:\Temp\@VES_AUTOGRN5.xml",0) #Const oUserFields = 152 (&H98)
write-host  $oUF5.add()  $cmp.GetLastErrorDescription()

$oUF6 = $cmp.GetBusinessObjectFromXML("C:\Temp\@VES_AUTOGRN6.xml",0) #Const oUserFields = 152 (&H98)
write-host  $oUF6.add()  $cmp.GetLastErrorDescription()