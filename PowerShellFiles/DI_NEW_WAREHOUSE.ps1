<#
  purpose: Copy old warehouse setting to New warehouse
           0. Duplicate old warehouse setting and give a new warehouse name
           1. query all itemcode,itemcost ,5 dimension  which were in old warehouse
           2. Add above item code into new warehouse
           3. Add G/L account determine rule for new warehouse 
           4. do stock revaluation for new warehouse 
          
  Date   : 2021.05.10

#>

$cmp = New-Object -COMobject "SAPbobsCOM.Company"
 


# 初始化值
$site = 'SZ'
$refered_warehosecode = 'SZ-B-SR'
$NewWarehouse = 'SZ-B-SG'
$ticknum = 'NEWWAREHOUSE'

 .  C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1 

 FN_CONNECTSAPB1 $CMP $SITE 

# Disable SP control 





# return BPGroup code
Function get_BPGrpCode($BPGrpName)
{
  $oBPGrp = $cmp.GetBusinessObject(10)  # oBusinessPartnerGroups
  $oRs = $cmp.GetBusinessObject(300)
  $oRs.DoQuery("Select groupcode from ocrg where GroupName = '$BPGrpName'")
  $oBPGrp.Browser.Recordset =$oRs
  if (!$oBPGrp.Browser.EoF){return $oBPGrp.Code}
}

# return ItemGroup number
Function get_ItmGrpCode($ItmGrpName)
{
 $oItmGrp = $cmp.GetBusinessObject(52) # oItemGroups
 $oRs = $cmp.GetBusinessObject(300)
 $ors.DoQuery("select ItmsGrpCod from oitb where itmsgrpnam = '$ItmGrpName'")
 $oItmgrp.Browser.Recordset = $oRs

 if (!$oItmGrp.Browser.EoF){return $oItmGrp.Number}

}

# list items and  price of old warehouse
$itemcodes = Import-Csv C:\Temp\Itemcode2.csv | %{ "'" + $_.itemcode + "'"} -ErrorAction Continue
$itemcodesString = $itemcodes -join ","  

# 如果上面item codes string不是空值
if ($itemcodesString)
{
    $sql_listItms = "select t0.ItemCode ,t0.whscode,'' linenum, t1.price ,'' OcrCode ,'' OcrCode2
    ,'' OcrCode3 ,''  OcrCode4, '' OcrCode5
     from OITw t0 inner join ITM1 t1 on t0.ItemCode = t1.ItemCode
    inner join  OITM t2 on t1.itemcode = t2.ItemCode left join oitw t00 on t0.itemcode = t00.itemcode and t00.whscode ='$($NewWarehouse)'
    where t0.WhsCode='$($refered_warehosecode)' and t1.PriceList =1 and t1.Price > 0  and t2.evalsystem <>'F'  and t2.InvntItem ='Y'
    AND  CASE WHEN  t2.frozenFor ='N' THEN  ISNULL(t2.validTo,GETDATE()+1) ELSE '19000101' END >GETDATE() and isnull(t00.avgprice,0)= 0 
    and  t0.itemcode in ($itemcodesString)
    "
}
else 
{
 $sql_listItms = "select t0.ItemCode ,t0.whscode,'' linenum, t1.price ,'' OcrCode ,'' OcrCode2
    ,'' OcrCode3 ,''  OcrCode4, '' OcrCode5
     from OITw t0 inner join ITM1 t1 on t0.ItemCode = t1.ItemCode
    inner join  OITM t2 on t1.itemcode = t2.ItemCode left join oitw t00 on t0.itemcode = t00.itemcode and t00.whscode ='$($NewWarehouse)'
    where t0.WhsCode='$($refered_warehosecode)' and t1.PriceList =1 and t1.Price > 0  and t2.evalsystem <>'F'  and t2.InvntItem ='Y'
    AND  CASE WHEN  t2.frozenFor ='N' THEN  ISNULL(t2.validTo,GETDATE()+1) ELSE '19000101' END >GETDATE() and isnull(t00.avgprice,0)= 0 
    "
}

# $sql_listItms | Out-File C:\Temp\tsql.txt

# list G/L account determine rule for old warehouse
$sql_GLRules = " SELECT 
t.AbsEntry
 ,t.PeriodCat
 ,t.RuleCode
 ,t.ItemCode,t.ItmsGrpCod,t.BPGrpCod,t.WhsCode,t.Comments,t.Active,t.FromDate,t.ToDate
 ,t.StockAct 'Stock Account'
 ,t.COGM_Act 'Cost of Goods Sold Account'
 ,t.AlocCstAct 'Allocation account'
 ,t.PricDifAct 'Price Difference Account'
 ,t.DfltIncom         'Revenue Account'
 ,t.DfltExpn  'Expense Account'
 ,t.VariancAct  'Variance Account'
 ,t.DfltLoss  'Stock Offset - Decr. Acct'
 ,t.DfltProfit  'Stock Offset - Incr. Acct'
 ,t.RturnngAct  'Sales Returns Account'
 ,t.ECIncome  'Revenue Account - EU'
 ,t.ECExepnses  'Expense Account - EU'
 ,t.ForgnIncm     'Revenue Account - Foreign'
 ,t.ForgnExpn         'Expense Account - Foreign'
 ,t.ExDiffAct  'Exchange Rate Differences Account'
 ,t.BalanceAct  'Goods Clearing Account'
 ,t.IncresGlAc  'G/L Increase Account'
 ,t.DecresGlAc   'G/L Decrease Account'
 ,t.WipAcct   'WIP Stock Account'
 ,t.WipVarAcct 'WIP Stock Variance Account'
 ,t.WipOffset       'WIP Offset P&L Account' 
 ,t.StockOffst       'Stock Offset P&L Account'    
 ,t.ExpClrAct       'Expense Clearing Account'      
 ,t.ARCMAct       'Sales Credit Account'        
 ,t.ARCMFrnAct       'Sales Credit Account - Foreign'  
 ,t.ARCMEUAct    'Sales Credit Account - EU'   
 ,t.APCMAct        'Purchase Credit Account'          
 ,t.APCMFrnAct          'Purchase Credit Acct - Foreign'            
 ,t.APCMEUAct 'Purchase Credit Account - EU'
 ,t.NegStckAct             'Negative Stock Adj. Acct'               
 ,t.ShpdGdsAct           'Shipped Goods Account'           
 ,t.VatRevAct           'VAT in Revenue Account'          
 ,t.StkInTnAct          'Stock in Transit Account'       
 ,t.StockRvAct    'Stock Revaluation Account'
,t.StkRvOfAct 'Stock Reval. Offset Acct'
  FROM OGAR t WHERE t.PeriodCat = CAST(year(getdate()) AS VARCHAR(4)) AND t.Active='Y' AND t.WhsCode ='$refered_warehosecode'"



  
# 在内存中处理数据
  $cmp.XMLAsString = 1
  $cmp.XmlExportType = 3
  
  $oRs = $cmp.GetBusinessObject(300)
  $oRs.DoQuery($sql_listItms)
  if (!$oRs.EoF) {[xml]$ItemCostXml=$oRs.GetAsXML()}

# 为item主数据 添加新仓库
  $oItm = $cmp.GetBusinessObject(4)
  $RowsNode = $ItemCostXml.SelectNodes('//row')
  $RowsCount = $RowsNode.Count

  # Loop over selected nodes
  $n = 1 
  foreach($node in $RowsNode){
    if ($oItm.GetByKey($node.ItemCode)){
       $oItm.WhsInfo.Add()
       $oItm.WhsInfo.WarehouseCode = $NewWarehouse
        Write-Host $cmp.CompanyDB " process $n/$RowsCount added item code " $node.ItemCode " into $NewWarehouse with error code : " $oItm.update() " description " $cmp.GetLastErrorDescription()
        $n+=1
    }
  }
  
   start-sleep 3

# 为 新仓库后者新物料组 设置 G/L acount determination rules .
 function AddGLRules ($absentry,$newRuleCode,$newWhse)                             # 此处参数 absentry 是获取已有的GL rule
 {   
  $oCmpService = $cmp.GetCompanyService()
  $oGLAccountAdvancedRulesService = $oCmpService.GetBusinessService(1470000057) # GLAccountAdvancedRulesService
  $oGLAccountAdvancedRule = $oGLAccountAdvancedRulesService.GetDataInterface(0) # glaarsGLAccountAdvancedRule
  $GLAccountAdvancedRuleParams  = $oGLAccountAdvancedRulesService.GetDataInterface(2) # glaarGLAccountAdvancedRuleParams

  $GLAccountAdvancedRuleParams.AbsoluteEntry = $absentry
       $oGLAccountAdvancedRule = $oGLAccountAdvancedRulesService.Get($GLAccountAdvancedRuleParams) 
       [XML]$Temp = $oGLAccountAdvancedRule.ToXMLString()
       $RuleCode = $Temp.SelectSingleNode("//Code")
       $ItemGroup = $Temp.SelectSingleNode("//ItemGroup")
       $Warehouse = $Temp.SelectSingleNode("//Warehouse")
        
        $RuleCode.InnerText = $newRuleCode                                            # assign new Rule code
       # $ItemGroup.InnerText = 343                                                   # assign new item group code
        $Warehouse.InnerText = $newWhse
      $oGLAccountAdvancedRule.FromXMLString($Temp.InnerXml)
      $oGLAccountAdvancedRulesService.Add($oGLAccountAdvancedRule)  |Out-Null
    
    Write-Host $cmp.CompanyDB " Add Acct Rule: " $newRuleCode " with error code:" $cmp.GetLastErrorCode() " and description is:" $cmp.GetLastErrorDescription()
 }

 $oRs_GLRules = $cmp.GetBusinessObject(300)
 $ors_glrules.DoQuery($sql_GLRules)                       # 查询 为 引用仓库，物料组所设置的 G/L advanced determination rules 

 if (!$oRs_GLRules.EoF)
  { 
 
      [xml]$xmlGLRules =  $ors_glrules.GetAsXML()         # 将查询结果保存到 XML 里面
       $RuleNodes = $xmlGLRules.SelectNodes("//row")      # 选择所有 row 
       
       $i = 1
       foreach($nd in $RuleNodes)                         # 遍历 row 里面的值
       {

         $newRulcd =   $NewWarehouse+$i                   # GL rule code 长度 20,注意检查。

         AddGLRules $nd.AbsEntry  $newRulcd $NewWarehouse    # 添加 新的 GL Acct rule
        
         $i+=1
   
       }

   }



 
  Start-Sleep 3


  $i=0                          # used to hashtable index ,it must start with 0  
  while ($i -lt $RowsCount ){
    $j = $i + 299               # define sub hashtable max index
    if ($j -lt $RowsCount) 
    {$300Rows = $RowsNode[$i..$j]}  # 一次读取哈希表中300行的数据，从0行 到定义的最大行，如果达到哈希表最大行，则取其行。
    else 
    { $lastrow= $RowsCount -1
    $300Rows = $RowsNode[$i..$lastrow]
    }

     $oMaterialRevaluation = $cmp.GetBusinessObject(162)
     $oMaterialRevaluation_line = $oMaterialRevaluation.Lines

     $oMaterialRevaluation.DocDate = Get-Date
     $oMaterialRevaluation.TaxDate = Get-Date
     $oMaterialRevaluation.RevalType = 'P'
     $oMaterialRevaluation.Comments = "intial item cost for new warehouse "

     # add 300 rows in one stock revaluation document
     $linenum = 0 
    foreach ($r in $300Rows){
           $oMaterialRevaluation_line.Add()
           $oMaterialRevaluation_line.SetCurrentLine($linenum)
           $oMaterialRevaluation_line.ItemCode = $r.ItemCode
           $oMaterialRevaluation_line.WarehouseCode = $NewWarehouse
           $oMaterialRevaluation_line.Price = [Double]$r.price
           
           # if there is no inventory of the new warehouse ,no need 5 dimension,but if has inventory, must define 5D
           # $oMaterialRevaluation_line.DistributionRule = 'CUSE'
           # $oMaterialRevaluation_line.DistributionRule2 = 'Local'
           # $oMaterialRevaluation_line.DistributionRule3 = 'EE'
           # $oMaterialRevaluation_line.DistributionRule4 = 'OTCO'
           # $oMaterialRevaluation_line.DistributionRule5 = 'FCGL'
         
           $linenum +=1
           
         }
    Write-Host "Stock revaluation Added with error: " $oMaterialRevaluation.Add() ” and description is:" $cmp.GetLastErrorDescription() 

    $i = $j + 1  # SUB hashtable next starting index
   
  }
    
    

  # Ensable SP control 
  $ticknum = ''

  $CMP.Disconnect()

Stop-Process -Id $PID


