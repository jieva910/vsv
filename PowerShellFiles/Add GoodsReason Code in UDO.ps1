
cls

$cmp = New-Object -ComObject 'sapbobscom.company'
$SourceSite    = "SZtst"

$ticktNum = 'add goods reason code'

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp $SourceSite


  fn_SAPB1_SP_control $ticktNum 'N' $SourceSite


$oCompServic = $cmp.GetCompanyService()
$oGeneralServic = $oCompServic.GetGeneralService('VES_TNMSGS')
  
# Add new code in existing UDO - [@VES_TNMSGS]
 $csvfiel =  Import-Csv C:\Temp\TN_CODE.csv
$oGeneralData = $oGeneralServic.GetDataInterface(1)  # GeneralData data interface

 foreach($row in $csvfiel)
 {      
  $oGeneralData.SetProperty("Code", $row.Code.trim())
  $oGeneralData.SetProperty("Name",$row.Name)
  $oGeneralData.SetProperty("U_VES_Name",$row.Name)
    $oGeneralData.SetProperty("U_VES_FrgnName",$row.Name)
  
  [void]$oGeneralServic.add($ogeneraldata); $cmp.GetLastErrorDescription()
 }

 
 # add VALUE IN UDT
  $udt = $cmp.GetBusinessObject(153) #Const oUserTables = 153 (&H99)
 $udt = $cmp.UserTables.Item('VES_REASONCODES')

  
  foreach($code in Import-Csv C:\Temp\GRCodelist.csv){
          $UDT.Code =$code.Code 
          $UDT.Name = $code.Name
         $UDT.add(); $CMP.GetLastErrorDescription()
 
         }
         
# Run SP via DI API
$oRs = $cmp.GetBusinessObject(300)
$oRs.DoQuery('C:\Temp\sql\TN_L_GoodsReturn.sql')
$oRs.DoQuery('C:\Temp\sql\TN_GoodsReturn.sql')

<#  Run below query in SBO_TN
  
  -- *** START of Goods Return Validation ***

IF @object_type = '21'

BEGIN
	EXEC VES_TN_GoodsReturn
		@TransType = @transaction_type,
		@IndexKey = @list_of_cols_val_tab_del,
		@ErrCode = @error OUTPUT,
		@ErrMsg = @error_message OUTPUT
END

#>


# ADD VALUE FOR UDO Special condition @VES_SPECCOND Table


$oCompServic = $cmp.GetCompanyService()
$oGeneralServic = $oCompServic.GetGeneralService('VES_SpecialCondition')
  
# Add new code in existing UDO
 $csvfiel =  Import-Csv C:\Temp\szconsum.csv -Delimiter ","
$oGeneralData = $oGeneralServic.GetDataInterface(1)  # GeneralData data interface

 foreach($row in $csvfiel)
 {      
  $oGeneralData.SetProperty("Code", $row.Code.trim())
  $oGeneralData.SetProperty("U_VesBPCode",$row.'Final User')
  $oGeneralData.SetProperty("U_VesItemCode",$row.ItemCode)
    $oGeneralData.SetProperty("U_VesSpecCondCd",$row.consuming)
  
  $oGeneralServic.add($ogeneraldata);$cmp.GetLastErrorDescription()

 }
