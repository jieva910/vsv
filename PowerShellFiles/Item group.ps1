
# add new item group then add neww GL accnt rules in each sapb1 site
# 2023 04 12

$ticktNum ="INC0346523"

$ItemGrpName = @(
'EXP-PF-HT SERVICES'
)

$csv = Import-Csv 'C:\Temp\new item group.csv'


 # get itemgroup code
Function get_itemgroupcode($itemgroupname)
  { 
     $oItmGrp = $cmp.GetBusinessObject(52)   #  oItemGroups
     $oRs = $cmp.GetBusinessObject(300)
    $oRs.DoQuery("select ItmsGrpCod from oitb where itmsgrpnam = '$itemgroupname'")
    $oItmGrp.Browser.Recordset = $oRs
    
    If ($oItmGrp.Browser.EOF -eq $False) { return $oItmGrp.Number}
}



# add new item group 
function fn_addItmGrp {
foreach( $itgrp in $ItemGrpName)
{
  
    $oItmGrp = $CMP.GetBusinessObject([sapbobscom.boobjecttypes]::oItemGroups)

    $oItmGrp.GroupName =  $itgrp
    $oItmGrp.UserFields.Fields.Item("U_Ves_AssetItem").value = 'N'
    $oItmGrp.UserFields.Fields.Item("U_Ves_CW").value =  'N'
    $oItmGrp.UserFields.Fields.Item("U_Ves_InvItem").value = 'N'
    $oItmGrp.UserFields.Fields.Item("U_Ves_SellItem").value =  'N'
    $oItmGrp.UserFields.Fields.Item("U_Ves_ManBatchNum").value =  'N'
    $oItmGrp.UserFields.Fields.Item("U_Ves_NoInvCons").value =  'N'
    $oItmGrp.UserFields.Fields.Item("U_Ves_OWN_FG").value =  'N'
    $oItmGrp.UserFields.Fields.Item("U_Ves_OWN_WIP").value =  'N'
    $oItmGrp.UserFields.Fields.Item("U_Ves_PrchseItem").value =  'N'
    $oItmGrp.UserFields.Fields.Item("U_Ves_RM_AX_PK").value =  'N'

    write-host $itgrp  $oItmGrp.Add()  $CMP.GetLastErrorDescription()

}
}


  #   add GL account rules

function fn_addGlRuls {
  foreach($row in $csv)
  {
   
      $oGLAccountAdvancedRulesService = $cmp.GetCompanyService().GetBusinessService(1470000057)  # GLAccountAdvancedRulesService
  $oGLAccountAdvancedRule = $oGLAccountAdvancedRulesService.GetDataInterface(0)               # glaarsGLAccountAdvancedRule
  $param = $oGLAccountAdvancedRulesService.GetDataInterface(2)   # glaarsGLAccountAdvancedRuleParams
  
      $oGLAccountAdvancedRule.Period = Get-Date -Format 'yyyy'
          
            $oGLAccountAdvancedRule.Code =  $row.'Advanced Rule Code'
          $oGLAccountAdvancedRule.ItemGroup = get_itemgroupcode $row.'Group Name'.TrimEnd()
               $oGLAccountAdvancedRule.ExpensesAccount = $row.'Expense Account' # Expense Account
                $oGLAccountAdvancedRule.ForeignExpensAcc = $row.'Expense Account - Foreign' # Expense Account - Foreign
                 $oGLAccountAdvancedRulesService.Add($oGLAccountAdvancedRule)
                   Write-Output ($row.'Advanced Rule Code' + " added with err code:" +  $cmp.GetLastErrorDescription())
    }
 }


  $sites = "kt"

  foreach($s in $sites)
  {
  
    
   $CMP = New-Object  -ComObject "SAPBOBSCOM.COMPANY"

  . C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

    FN_CONNECTSAPB1 $CMP $s 
   Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"

    fn_SAPB1_SP_control $ticktNum 'N' $s
  
  
    fn_addItmGrp
    
    Start-Sleep -Seconds 3


    fn_addGlRuls
  
  
  
    #Enable sapb1 TN SP control
    Write-Host -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
    $ticktNum2 =''
    fn_SAPB1_SP_control $ticktNum2 'Y' $s
    Release-Ref $cmp
  
  }