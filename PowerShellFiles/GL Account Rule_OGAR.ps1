
<#
  Purpose : Add G/L account determination rule for EMEA sapb1 (SR,LA,GE,GH,OS,TC)
  Date    : 2020.12
#>


cls
$cmp = new-object -ComObject "sapbobscom.company"
$SourceSite    = "wn"
$ticknum = 'INC0283097'
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp $SourceSite

# Disable SP control 
fn_SAPB1_SP_control $ticknum 'N'  $SourceSite 

# Update GL account determination rule 
  $oRs = $cmp.GetBusinessObject(300)
  $oGLAccountAdvancedRulesService = $cmp.GetCompanyService().GetBusinessService(1470000057)  # GLAccountAdvancedRulesService
  $oGLAccountAdvancedRule = $oGLAccountAdvancedRulesService.GetDataInterface(0)               # glaarsGLAccountAdvancedRule
  $param = $oGLAccountAdvancedRulesService.GetDataInterface(2)   # glaarsGLAccountAdvancedRuleParams
  
  $csv = Import-Csv C:\Temp\ASACCTRULES.csv

  foreach($row in $csv)
  {
   
    $sql = "select t.AbsEntry from ogar t  where t.Year = year(getdate()) and t.RuleCode = '$($row.'Rule code')'"  
    $oRs.DoQuery($sql)
    if (!$oRs.EoF)
    { 
     $param.AbsoluteEntry = $oRs.Fields.Item(0).value
     
      $oGLAccountAdvancedRule = $oGLAccountAdvancedRulesService.Get($param)
         $oGLAccountAdvancedRule.ExchangeRateDifferencesAcct = [string]$row.'Exchange Rate Differences Account'
      <# $oGLAccountAdvancedRule.RevenuesAccount =$row.'Revenue Account'
        $oGLAccountAdvancedRule.InventoryAccount=$row.'Stock Account'
         $oGLAccountAdvancedRule.CostAccount =$row.'Cost of Goods Sold Account'
          $oGLAccountAdvancedRule.VarienceAccount = $row.'Variance Account'
           $oGLAccountAdvancedRule.PriceDifferenceAcc=$row.'Price Difference Account'
            $oGLAccountAdvancedRule.DecreasingAccount = $row.'Stock Offset - Decr. Acct'
            $oGLAccountAdvancedRule.IncreasingAccount = $row.'Stock Offset - Incr. Acct' # Stock Offset - Incr. Acct
             $oGLAccountAdvancedRule.ReturningAccount = $row.'Sales Returns Account' # Sales Returns Account
             $oGLAccountAdvancedRule.ForeignRevenueAcc = $row.'Revenue Account - Foreign' # Revenue Account - Foreign #>
        # Exchange Rate Differences Account
         <# $oGLAccountAdvancedRule.StockInTransitAccount = $row.'Stock in Transit Account' # Stock in Transit Account
      $oGLAccountAdvancedRule.GLIncreaseAcct = $row.'G/L Increase Account'
      $oGLAccountAdvancedRule.GLDecreaseAcct = $row.'G/L Decrease Account'
       $oGLAccountAdvancedRule.SalesCreditAcc = $row.'Sales Credit Account' # Sales Credit Account
     $oGLAccountAdvancedRule.SalesCreditForeignAcc = $row.'Sales Credit Account - Foreign' # Sales Credit Account - Foreign   #>    
      $oGLAccountAdvancedRulesService.Update($oGLAccountAdvancedRule) 
      Write-Output ($row.'Rule code' + " update with err code:" +  $cmp.GetLastErrorCode())

    }
  }
  
 

 
 # get itemgroup code
Function get_itemgroupcode($itemgroupname)
  { 
     $oItmGrp = $cmp.GetBusinessObject(52)   #  oItemGroups
     $oRs = $cmp.GetBusinessObject(300)
    $oRs.DoQuery("select ItmsGrpCod from oitb where itmsgrpnam = '$itemgroupname'")
    $oItmGrp.Browser.Recordset = $oRs
    
    If ($oItmGrp.Browser.EOF -eq $False) { return $oItmGrp.Number}
}

 # get bp group code
Function get_bpgroup($bpgroupname)
 {  
     $obpgrp = $cmp.GetBusinessObject(10)   # oBusinessPartnerGroups
     $oRs = $cmp.GetBusinessObject(300)
    $oRs.DoQuery("Select groupcode from ocrg where GroupName = '$bpgroupname'")
    $obpgrp.Browser.Recordset = $oRs
    
    If ($obpgrp.Browser.EOF -eq $False) 
      { return $obpgrp.code }
 }



# Add_GlAcctRule
  $GlRules_csv = Import-Csv C:\Temp\GLAcctRule\wnnewitemgrup.csv
 
  foreach($row in $GlRules_csv)
  {   

  $oGLAccountAdvancedRulesService = $cmp.GetCompanyService().GetBusinessService(1470000057)  # GLAccountAdvancedRulesService
  $oGLAccountAdvancedRule = $oGLAccountAdvancedRulesService.GetDataInterface(0)               # glaarsGLAccountAdvancedRule
  $param = $oGLAccountAdvancedRulesService.GetDataInterface(2)   # glaarsGLAccountAdvancedRuleParams
  
      $oGLAccountAdvancedRule.Period = Get-Date -Format 'yyyy'
          
            $oGLAccountAdvancedRule.Code =  $row.'Rule code'
            IF( $row.'Item Code' -ne 'All'){$oGLAccountAdvancedRule.ItemCode = $row.'Item Code'}
            IF( $row.'Item Group' -ne 'All'){$oGLAccountAdvancedRule.ItemGroup = get_itemgroupcode $row.'Item Group'.TrimEnd()}
             IF( $row.'Business Partner Group' -ne 'All'){$oGLAccountAdvancedRule.BPGroup = get_bpgroup $row.'Business Partner Group'.TrimEnd()}
             IF( $row.'Warehouse Code' -ne 'All'){$oGLAccountAdvancedRule.Warehouse = $row.'Warehouse Code'}
            $oGLAccountAdvancedRule.InventoryAccount = $row.'Stock Account' # Stock Account
            $oGLAccountAdvancedRule.CostAccount = $row.'Cost of Goods Sold Account' # Cost of Goods Sold Account
            $oGLAccountAdvancedRule.TransferAccount = $row.'Allocation Account' # Allocation account
            $oGLAccountAdvancedRule.PriceDifferenceAcc = $row.'Price Difference Account' # Price Difference Account
            $oGLAccountAdvancedRule.RevenuesAccount = $row.'Revenue Account' # Revenue Account
              $oGLAccountAdvancedRule.ExpensesAccount = $row.'Expense Account' # Expense Account
              $oGLAccountAdvancedRule.VarienceAccount = $row.'Variance Account' # Variance Account
            $oGLAccountAdvancedRule.DecreasingAccount = $row.'Stock Offset - Decr. Acct' # Stock Offset - Decr. Acct
            $oGLAccountAdvancedRule.IncreasingAccount = $row.'Stock Offset - Incr. Acct' # Stock Offset - Incr. Acct
            $oGLAccountAdvancedRule.ReturningAccount = $row.'Sales Returns Account' # Sales Returns Account
           # $oGLAccountAdvancedRule.EURevenuesAccount = $row.'Revenue Account - EU' # Revenue Account - EU
           # $oGLAccountAdvancedRule.EUExpensesAccount = $row.'Expense Account - EU' # Expense Account - EU
             $oGLAccountAdvancedRule.ForeignRevenueAcc = $row.'Revenue Account - Foreign' # Revenue Account - Foreign
            $oGLAccountAdvancedRule.ForeignExpensAcc = $row.'Expense Account - Foreign' # Expense Account - Foreign
            $oGLAccountAdvancedRule.ExchangeRateDifferencesAcct = $row.'Exchange Rate Differences Account' # Exchange Rate Differences Account
            $oGLAccountAdvancedRule.GoodsClearingAcct = $row.'Goods Clearing Account' # Goods Clearing Account
            $oGLAccountAdvancedRule.GLIncreaseAcct = $row.'G/L Increase Account' # G/L Increase Account
             $oGLAccountAdvancedRule.GLDecreaseAcct = $row.'G/L Decrease Account' # G/L Decrease Account
             $oGLAccountAdvancedRule.WipAccount = $row.'WIP Stock Account' # WIP Stock Account
             $oGLAccountAdvancedRule.WipVarianceAccount = $row.'WIP Stock Variance Account' # WIP Stock Variance Account
             $oGLAccountAdvancedRule.WipOffsetProfitAndLossAccount = $row.'WIP Offset P&L Account' # WIP Offset P&L Account
             $oGLAccountAdvancedRule.InventoryOffsetProfitAndLossAccount = $row.'Stock Offset P&L Account' # Stock Offset P&L Account
             $oGLAccountAdvancedRule.ExpenseClearingAct = $row.'Expense Clearing Account' # Expense Clearing Account
             $oGLAccountAdvancedRule.SalesCreditAcc = $row.'Sales Credit Account' # Sales Credit Account
              $oGLAccountAdvancedRule.SalesCreditForeignAcc = $row.'Sales Credit Account - Foreign' # Sales Credit Account - Foreign
             # $oGLAccountAdvancedRule.SalesCreditEUAcc = $row.'Sales Credit Account - EU' # sales Credit Account - EU
               $oGLAccountAdvancedRule.PurchaseCreditAcc = $row.'Purchase Credit Account' # Purchase Credit Account
             #  $oGLAccountAdvancedRule.EUPurchaseCreditAcc = $row.'Purchase Credit Account - EU' # Purchase Credit Account - EU
               $oGLAccountAdvancedRule.ForeignPurchaseCreditAcc = $row.'Purchase Credit Acct - Foreign' # Purchase Credit Acct - Foreign
             $oGLAccountAdvancedRule.NegativeInventoryAdjustmentAccount = $row.'Negative Stock Adj. Acct' # Negative Stock Adj. Acct
                $oGLAccountAdvancedRule.ShippedGoodsAccount = $row.'Shipped Goods Account' # Shipped Goods Account
                $oGLAccountAdvancedRule.VATInRevenueAccount = $row.'VAT in Revenue Account' # VAT in Revenue Account
                $oGLAccountAdvancedRule.StockInTransitAccount = $row.'Stock in Transit Account' # Stock in Transit Account 

    $oGLAccountAdvancedRulesService.Add($oGLAccountAdvancedRule)
     Write-Output ($row.'Rule code' + " added with err code:" +  $cmp.GetLastErrorDescription())

        Release-Ref($oGLAccountAdvancedRulesService)
   }

          
             
    

    Function Savetoxml()
  { 
    
    $param = $oGLAccountAdvancedRulesService.GetDataInterface(2)   # glaarsGLAccountAdvancedRuleParams
      $param.AbsoluteEntry = 12397
       $oGLAccountAdvancedRule = $oGLAccountAdvancedRulesService.Get($param)
 
    $oGLAccountAdvancedRule.ToXMLFile("C:\temp\GLACCT.XML")
   }  


    # delete GL rule
    $param = $oGLAccountAdvancedRulesService.GetDataInterface(2)   # glaarsGLAccountAdvancedRuleParams
      $param.AbsoluteEntry = 5577
     $oGLAccountAdvancedRulesService.Delete($param)




# UPDATE Single acct rule ################################

 $oGLAccountAdvancedRulesService = $cmp.GetCompanyService().GetBusinessService(1470000057)  # GLAccountAdvancedRulesService
  $oGLAccountAdvancedRule = $oGLAccountAdvancedRulesService.GetDataInterface(0)               # glaarsGLAccountAdvancedRule
  $param = $oGLAccountAdvancedRulesService.GetDataInterface(2)   # glaarsGLAccountAdvancedRuleParams
  
     $oRs = $cmp.GetBusinessObject(300)
     $sql = "select t.AbsEntry from ogar t  where t.Year = year(getdate()) and t.RuleCode = 'ALL MISEMI82'"  
    $oRs.DoQuery($sql)

     if (!$oRs.EoF){
     $param.AbsoluteEntry = $oRs.Fields.Item(0).value
     $oGLAccountAdvancedRule = $oGLAccountAdvancedRulesService.Get($param)

    # $oGLAccountAdvancedRule.ItemCode ='XY03612'

          $oGLAccountAdvancedRule.DecreasingAccount = '550201-01-02' # Stock Offset - Decr. Acct
            $oGLAccountAdvancedRule.IncreasingAccount = '550201-01-02' # Stock Offset - Incr. Acct
               $oGLAccountAdvancedRule.GLIncreaseAcct = '550201-03-01' # G/L Increase Account
             $oGLAccountAdvancedRule.GLDecreaseAcct = '550201-03-01' # G/L Decrease Account
        $oGLAccountAdvancedRulesService.Update($oGLAccountAdvancedRule) 
    $cmp.GetLastErrorCode()
    }