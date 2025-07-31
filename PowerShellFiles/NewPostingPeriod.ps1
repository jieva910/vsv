cls
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$SourceSite    = "kttst"
# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite

#get company service
$oCompanyService = $cmp.GetCompanyService()
# ++++++++++++++++++++++ add posting period ,test successfully
   
     $oPeriodCategory =$oCompanyService.GetDataInterface(2) 
     $PeriodCategoryParams	 =$oCompanyService.GetDataInterface(3) 
    # $FinancePeriodParams   = $oCompanyService.GetDataInterface(6)
    # $FinancePeriodParams.PeriodIndicator = '2022'

    $oPeriodCategory.PeriodName = "2022"
    $oPeriodCategory.PeriodCategory  = '2022'   # for GL Acct determine rule 
    # set the period type can be year,quater,month or day
    # (e.g. spt_Year=0,spt_quater=1,spt_month=2,spt_days)
    $oPeriodCategory.SubPeriodType = 2
    $oPeriodCategory.NumberOfPeriods = 12
    $oPeriodCategory.FinancialYear = '2022'
    # set the beginning of Financial Year
    $oPeriodCategory.BeginningofFinancialYear ="2022-01-01"
    $newEntry = $oCompanyService.CreatePeriod($oPeriodCategory)

   $oFinancePeriods = $oCompanyService.GetFinancePeriods($newEntry)

   foreach($oFinancePed in $oFinancePeriods)
   {
     
    $oFinancePed.TaxDateFrom ='2020-01-01'
    $oFinancePed.TaxDateTo ='2022-12-31'
    $oFinancePed.ValueDateFrom ='2022-01-01'
    $oFinancePed.ValueDateTo = '2023-12-31'
    $oCompanyService.UpdateFinancePeriod($oFinancePed)

   }
# ++++++++++++++++++++++ add posting period 