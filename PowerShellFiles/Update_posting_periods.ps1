$cmp = New-Object -ComObject 'sapbobscom.company'



$cmp.Server = 'SZ-sapstg91'
$cmp.CompanyDB ='SAPB1_YK_TST'
$cmp.DbServerType = 10
$cmp.UserName ='\'
$cmp.Password =''
# $cmp.DbUserName=$cmpdbuser
# $cmp.DbPassword=$cmpdbpwd
$cmp.UseTrusted=$true
$cmp.LicenseServer = 'SZ-TSTSAPLIC92'


[void]$cmp.Connect()

if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;break
}
else { Write-Host -ForegroundColor Cyan $cmp.CompanyDB connected successfully}



# get company service
$oCompanyService = $cmp.GetCompanyService()

# get Period Category Collection
$oPeriodCategoryColl = $oCompanyService.GetPeriods()

# print all periods
For ( $i = 0 ;$i -lt $oPeriodCategoryColl.Count ;$i++)
{   # get period category
    $oPerCategory = $oCompanyService.GetPeriod($oPeriodCategoryColl.Item($i))
  
      # $oPerCategory.FinancialYear -eq '2014'
     # $oPerCategory.PeriodName

    # get all finance periods (if the sub period isn# t a year then it
    # has more than one finance period, for example sub period month has 12 finance periods)
    $oFinancePeriods = $oCompanyService.GetFinancePeriods($oPeriodCategoryColl.Item($i))

    For ($j = 0 ; $j -lt $oFinancePeriods.Count ;$j++)
     {   # get finance period
        $oFinancePeriod = $oFinancePeriods.Item($j)
        # print the period name
        $oFinancePeriod.PeriodName
        $oFinancePeriod.PeriodStatus = 2      #ltPeriodClosing
        $oCompanyService.UpdateFinancePeriod($oFinancePeriod)
         $cmp.GetLastErrorDescription()
   }

}
