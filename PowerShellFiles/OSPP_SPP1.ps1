  


  
cls

$SourceSite    = "sztst"

$ticktNum = 'add goods reason code'

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite

fn_SAPB1_SP_control $ticktNum 'N' $SourceSite
  
$csv = Import-csv C:\Temp\spp1.csv -Delimiter ','

foreach($r in $csv )
{
      $oSp = $cmp.GetBusinessObject(7)

      [datetime]$newDateFrm=Get-Date $r.Datefrom  -Format "yyyy-MM-dd" 
      [datetime]$newDateto=Get-Date $r.dateto  -Format "yyyy-MM-dd"
      
        # Update price which exists
      If ($oSp.GetByKey($r.Itemcode, $r.Cardcode) )
        {
          $linenum = $oSp.SpecialPricesDataAreas.Count
         
          If ($linenum -gt 0)  #if has historical price, then update last price dateto as new datefrom -1
            { 
              [DateTime]$spDateFrom=Get-Date $oSp.SpecialPricesDataAreas.DateFrom  -Format "yyyy-MM-dd" 
              $compareDate = $spDateFrom.CompareTo( $newDateFrm)
               If ($compareDate -eq 1) { write-host $r.Itemcode " Last historical price Date from is "  $spDateFrom   "while new price date from is " $newDateFrm ", please check and update it manually"
                        continue  }  # compare date and quit loop

             $lastlinenum = $linenum - 1
             $oSp.SpecialPricesDataAreas.SetCurrentLine($lastlinenum) #compare last price date from to current date from         
             $lastenddate=$newDateFrm.AddDays(-1)
             $oSp.SpecialPricesDataAreas.Dateto = $lastenddate
          
            }
         
           [VOID]$oSp.SpecialPricesDataAreas.Add()
           $oSp.SpecialPricesDataAreas.DateFrom = $newDateFrm
           $oSp.SpecialPricesDataAreas.Dateto = $newDateto
           $oSp.SpecialPricesDataAreas.SpecialPrice =  $r.price
           $oSp.SpecialPricesDataAreas.PriceCurrency = 'RMB'
         $lRetCode = $oSp.Update()
         write-host $r.Itemcode " Updated with error code :" $lRetCode " and error description is " $cmp.GetLastErrorDescription()
        } 
       Else # Add new item price 
        {
         
         $oSp.CardCode =  $r.cardcode
         $oSp.PriceListNum = 0
         $oSp.ItemCode =$r.itemcode
         $oSp.Price =$r.price
     
         $oSp.SpecialPricesDataAreas.PriceListNo = 0
         $oSp.SpecialPricesDataAreas.DateFrom =$newDateFrm
         $oSp.SpecialPricesDataAreas.Dateto = $newDateto
         $oSp.SpecialPricesDataAreas.Discount = 0
         $oSp.SpecialPricesDataAreas.SpecialPrice = $r.price
            $oSp.SpecialPricesDataAreas.PriceCurrency = 'RMB'
          $lRetCode = $oSp.Add()
         Write-Host $r.Itemcode " Added with error code :"$lRetCode " and error description is "$cmp.GetLastErrorDescription()
         }
       
}

