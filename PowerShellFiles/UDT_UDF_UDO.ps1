
 




# 2022.10.
# supplier consignment process new 
$cmp = New-Object -ComObject 'sapbobscom.company'

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$site = "sztst"

 Fn_ConnectSAPB1 $cmp  $site

 $cmp.XmlExportType  = 3  
 $cmp.XMLAsString = 0 

# 1. export UDT #########################################################################################################
  $oUT = $cmp.GetBusinessObject(153) #Const oUserTables = 153 (&H99)
 
 if ( $out.GetByKey('VES_CONSGNhdr')) { $oUT.SaveToFile('c:\temp\VES_CONSGNhdr.xml')  }
 
 if ( $out.GetByKey('VES_CONSGNDTL')) { $oUT.SaveToFile("c:\temp\VES_CONSGNDTL.xml")  }

 # Export User  fileds 

 $oUF  = $cmp.GetBusinessObject(152)
$sql_UserFields = "SELECT FieldID,AliasID from cufd where tableid = '@VES_CONSGNhdr'" 
  
  $ors  = $cmp.GetBusinessObject(300)
  $ors.doquery("SELECT FieldID,AliasID from cufd where tableid = '@VES_CONSGNhdr'" )

  $xml = [xml]$ors.getasxml()

  $nodes = $xml.selectnodes("//row")

  foreach($n in $nodes)
  {
     if ( $oUF.GetByKey('@VES_CONSGNhdr',$n.FieldID)) { $oUF.SaveXML("c:\temp\$($n.AliasID).xml")  }

    }


 
 # export UDO
 $udo = $cmp.GetBusinessObject(206)
  if ( $udo.GetByKey('VES_CONSGN')) { $udo.SaveXML('C:\Temp\VES_CONSGN.xml')  }

# 2. Import udt,udf,udo ################################################################################################

 # import User Table 1
 $udtspath = "C:\Temp\SupplierConsignment\UDT_UDF_UDO\1.UDT\"
 $udtsXMLfiles = Get-ChildItem -Path "$udtspath*.xml" 

 foreach ($xmlfile in $udtsXMLfiles){
$oUT1=$cmp.GetBusinessObjectFromXML("$udtspath$($xmlfile.Name)",0)
 $oUT1=$cmp.GetBusinessObjectFromXML("C:\Temp\SupplierConsignment\UDT_UDF_UDO\1.UDT\VES_CONSGNhdr.xml",0)
Write-Host  "Add user Table with error code: "   $oUT1.Add()  " and error description: "  $cmp.GetLastErrorDescription()
  Release-Ref ($oUT1)
  }
 
 
# import User Fields

  $udfspath = "C:\Temp\SupplierConsignment\UDT_UDF_UDO\2.UDF\"
 $udfsXMLfiles = Get-ChildItem -Path "$udfspath*.xml" 

 foreach($f in $udfsXMLfiles){
 $oUF1 = $cmp.GetBusinessObjectFromXML("$udfspath$($f.Name)",0) #Const oUserFields = 152 (&H98)
write-host "Add UDFs with error code: " $oUF1.add() " and error description: " $cmp.GetLastErrorDescription()
Release-Ref ($oUF1)
  }


# Import UDO 
 $udoXMLfiles = Get-ChildItem -Path "C:\Temp\SupplierConsignment\UDT_UDF_UDO\3.UDO\VES_CONSIGN.xml" 
$udo = $cmp.GetBusinessObject(206)   
$udo.Browser.ReadXml($udoXMLfiles,0)


write-host "Add UDO with error code: " $udo.add() " and error description: " $cmp.GetLastErrorDescription()  