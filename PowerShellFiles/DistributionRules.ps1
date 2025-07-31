  
cls

$SourceSite    = "CNBGT"



$cmp = new-object -ComObject "sapbobscom.company"

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
 Fn_ConnectSAPB1 $cmp  $SourceSite




$oCmpSrv = $cmp.GetCompanyService()
$oDLservice  = $oCmpSrv.GetBusinessService(62)  # ServiceTypes.DistributionRulesService
$oDLParams = $oDLservice.GetDataInterface(2)    # drsDistributionRuleParams

# Get distribution rule
$oDLParams.FactorCode = "SZG-XTB"
Try
   { $oDL = $oDLservice.GetDistributionRule($oDLParams) }
Catch 

            { $ex = $_.exception }




# Update distribution rule  ,注意 DR Lines count 值
# $oDL = $oDLservice.GetDataInterface(0)   #drsDistributionRule
$oDL.FactorCode = "SZSS_GA4"
#$oDL.FactorDescription = "SZG-HDCNC1"
$oDL.InWhichDimension = 5
$oDL.TotalFactor = 100

$newRow = $oDL.DistributionRuleLines.Count 
$oDL.DistributionRuleLines.Add()
$oDL.DistributionRuleLines.Item($newRow).CenterCode = "SLGA"
$oDL.DistributionRuleLines.Item($newRow).TotalInCenter = "50"
$oDL.DistributionRuleLines.Item($newRow).Effectivefrom = '2022-12-01'
$oDL.DistributionRuleLines.Item($newRow).EffectiveTo = '2022-12-30'
$oDL.DistributionRuleLines.Add()
$oDL.DistributionRuleLines.Item($newRow+1).CenterCode = "VISO"
$oDL.DistributionRuleLines.Item($newRow+1).TotalInCenter = "50"
$oDL.DistributionRuleLines.Item($newRow+1).Effectivefrom = '2022-12-01'
$oDL.DistributionRuleLines.Item($newRow+1).EffectiveTo = '2022-12-30'

Try
 {  $oDLservice.UpdateDistributionRule($oDL) }
Catch 
  {$_.exception}



                # 尝试使用XML,测试过了，行不通。2022.08.17

                     $CMP.XmlExportType = 3 
                    $oCmpSrv = $cmp.GetCompanyService()
                    $oDLservice  = $oCmpSrv.GetBusinessService(62)  #ServiceTypes.DistributionRulesService

                    # Get distribution rule
                    $oDLParams = $oDLservice.GetDataInterface(2)  # drsDistributionRuleParams
                    $oDLParams.FactorCode = "SZGHDCN1"
                    Try
                       { $oDL = $oDLservice.GetDistributionRule($oDLParams) }
                    Catch 
                        { $_.exception }

                    $oDL = $oDLservice.GetDataInterface(0)   #drsDistributionRule
                    $oDL.ToXMLFile('C:\TEMP\DR.XML')


# 以下是批量更新 DR,成功在Budget DB 里面运行 。2024.8.26

   
    $csv =import-csv C:\Temp\DistrubitionRule\2024BUDGET3.csv -Delimiter ","
    $GRP = $CSV |  Group-Object -Property "AllocationKey"
    $oCmpSrv = $cmp.GetCompanyService()
    $oDLservice  = $oCmpSrv.GetBusinessService(62)  # ServiceTypes.DistributionRulesService
   $g = $null
    foreach($g in  $GRP)                         # 按allocationkey 分组，根据group name 主键，然后遍历Group里面的值
    { 
      $oDLParams = $oDLservice.GetDataInterface(2)    # drsDistributionRuleParams
      $DRCode = $g.Name
      $ex = $null 
       $oDLParams.FactorCode = $DRCode
        Try                                             
           { $oDL = $oDLservice.GetDistributionRule($oDLParams) }
        Catch 
            { $ex = $_.exception }

        if ($ex.ErrorCode -eq -2028 )       # 如果 no matching record found,则 新增,否则 更新
           {
             $oDL_new = $oDLservice.GetDataInterface(0)   #drsDistributionRule
             $oDL_new.FactorCode = $DRCode
             $oDL_new.FactorDescription = $DRCode
            
               $i = 0                        # 每组里面的循环次数
               $gp = $null
            foreach( $gp in $g.Group)
                  {
                      $oDL_new.InWhichDimension = $gp.AllocationDim
                      $oDL_new.TotalFactor = $gp.OcrTotal
                       $oDL_new.UserFields.Item("U_Ves_DefCostCentre").value = $gp.TransCode  # notice --> UserFields.Fields("U_VES_FrgnE_MailBody").value
                       $oDL_new.UserFields.Item("U_Ves_ContrDim1").value = $gp.Dim1
                       $oDL_new.UserFields.Item("U_Ves_ContrDim4").value = $gp.Dim4
                       $oDL_new.UserFields.Item("U_VES_ContrDim5").value = $gp.Dim5
                       $oDL_new.UserFields.Item("U_Ves_AcctType").value = $gp.AcctType
                       $oDL_new.UserFields.Item("U_Ves_Account").value = $gp.AccountCode
                        $oDL_new.UserFields.Item("U_VES_Branch").value = $gp.Company
                         $oDL_new.UserFields.Item("U_Ves_ContrDim2").value =$gp.Dim2                      

                    $oDL_new.DistributionRuleLines.Add()
                    $oDL_new.DistributionRuleLines.Item($i).CenterCode = $gp.PrcCode
                    $oDL_new.DistributionRuleLines.Item($i).TotalInCenter = [double]$gp.'BU'
                    $oDL_new.DistributionRuleLines.Item($i).Effectivefrom = '2024-08-01'
                    $oDL_new.DistributionRuleLines.Item($i).EffectiveTo = '2024-08-31'
    
                      $i+=1
                    }
  
   
               Try
                     {  $oDLservice.AddDistributionRule($oDL_new) | Out-Null
                        Write-Host $DRCode " added with error code:" $cmp.GetLastErrorCode()  }
               Catch 
                      { Write-Host $DRCode $_.exception }
           }
        else 
           {   
                    $oDL.FactorCode = $DRCode
                   # $oDL.FactorDescription = 
                    $newRow = $oDL.DistributionRuleLines.Count  # 获得当前Distribution rule 行数

                    $i = 0              # 每组里面的循环次数
                    $gp = $null
                  foreach( $gp in $g.Group)
                  {
                    $oDL.InWhichDimension = $gp.AllocationDim
                    $oDL.TotalFactor = [double]$gp.OcrTotal
                    $oDL.DistributionRuleLines.Add()
                    $oDL.DistributionRuleLines.Item($newRow+$i).CenterCode = $gp.PrcCode
                    $oDL.DistributionRuleLines.Item($newRow+$i).TotalInCenter = [double]$gp.'BU'
                    $oDL.DistributionRuleLines.Item($newRow+$i).Effectivefrom = '2024-08-01'
                    $oDL.DistributionRuleLines.Item($newRow+$i).EffectiveTo = '2024-08-31'
    
                      $i+=1
                    }
  
   
                    Try
                     {  $oDLservice.UpdateDistributionRule($oDL) | Out-Null
                        Write-Host $DRCode " updated with error code:" $cmp.GetLastErrorCode()  }
                    Catch 
                      { Write-Host $DRCode $_.exception }

                }
           
      }