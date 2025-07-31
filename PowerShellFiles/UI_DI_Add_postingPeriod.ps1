
cls
$sitelist = "SAPB1_kt_TST","SAPB1_sz_TST"
$ticktNum ="INC0291371"


$newpostingperid = [string]((get-date).AddYears(1).Year)

$find_new_posting_periods_sql = "SELECT count(code) FROM OFPR WHERE Code LIKE '%$($newpostingperid)%'"


$oCompany = new-object -ComObject "SAPBOBSCOM.Company"

# Connect to SBO via UI API
Function SetApplication {
              $SboGuiApi = New-Object -comobject "SAPbouiCOM.SboGuiApi"
              $sConnectionString =  "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

              $SboGuiApi.Connect($sConnectionString)
              $SboGuiApi.GetApplication()
          
        }
 #Connect with connection string
$SBO_Application = SetApplication
  function SetConnectionContext {
    
             $sCookie = $oCompany.GetContextCookie()
             $sConnectionContext = $SBO_Application.Company.GetConnectionContext($sCookie)
             If ($oCompany.Connected ){$oCompany.Disconnect()}
             return $oCompany.SetSboLoginContext($sConnectionContext)
        }

    # Connect to SBO via DI API

    Function ConnectToCompany {
       Return $oCompany.Connect()
    }




# Control the SP of blocking Super user
Function fn_SAPB1_SP_control($ticknum,$YesNO,$site)
  {   
   $oCompServic = $oCompany.GetCompanyService()
   $oGeneralServic = $oCompServic.GetGeneralService('VES_TNMSGS')
   $oGeneralParams = $oGeneralServic.GetDataInterface(3)  # gsGeneralDataParams = 3
   $oGeneralParams.SetProperty('Code','9999999')
   $oGeneralData = $oGeneralServic.GetByParams($oGeneralParams)
   
   # check current SP control status
   $isActive =  $oGeneralData.GetProperty('U_VES_Active')
   $comment = $ticknum
   if ($isActive -ne $YesNO)
     {
          switch ($site)
           {  { $site -in "xx"} {$ogeneraldata.SetProperty('U_VES_COmments',$comment)} 
     
              Default { $ogeneraldata.SetProperty('U_VES_Comments',$comment)}
           }
           $ogeneraldata.SetProperty('U_VES_Active',$YesNO)
           $oGeneralServic.Update($ogeneraldata)

      }
}




 foreach($site in $sitelist)                # 循环登录不同的 COMPANY DB
{
  $SBO_Application.ActivateMenuItem("3329") #open choose company list

  $oForm = $SBO_Application.Forms.ActiveForm
    $oItem = $oForm.Items.Item("4")
    $Omatrix = $oItem.Specific
    $oChkbox = $oForm.Items.Item("1470000128").Specific
    $strDBSvr = $oForm.Items.Item("420000125").Specific.Value
   Foreach ( $j in  1..$Omatrix.VisualRowCount)
    { 
      $strCompname = $Omatrix.Columns.Item("2").Cells($j).Specific.Value
    If ($strCompname -eq $site  ) {
       If (!$oChkbox.Checked) { $oChkbox.Item.Click(0)}
       $Omatrix.Columns.Item("2").Cells($j).Click(1)  #log on the company db with windows domain account
      
       #break                                          # 成功登录之后，退出当前循环
       
       Start-Sleep 5                                  # 开始   connect to DI 

        if (SetConnectionContext -ne 0 ) {$SBO_Application.MessageBox("Failed setting a connection to DI API");Exit}
        if (ConnectToCompany -ne 0 ) {$SBO_Application.MessageBox("Failed connecting to the company's Data Base") ; Exit}
        Write-Host -BackgroundColor Cyan "DI Connected To: " $oCompany.CompanyName

         Start-Sleep 10

         $ors  = $oCompany.GetBusinessObject(300)

         $ors.doquery($find_new_posting_periods_sql)
         $ors.Fields.Item(0).value
    if (!$ors.EoF -and  $ors.Fields.Item(0).value -gt 0 ) { Write-Host "already had new posting periods $($newpostingperid) " ; break }
        else {  

              #disable SP contorl in TNMSGS UDO
             Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"
             fn_SAPB1_SP_control $ticktNum 'N'

           $SBO_Application.ActivateMenuItem("8210")                                                  # Open 期间标识

           $oPeriodindicators =  $SBO_Application.Forms.ActiveForm
  
    $oPIRows = $oPeriodindicators.Items.Item(“3”).specific
     $lastrow = $oPIRows.RowCount
      $oPIRows.Columns.Item("Indicator").Cells( $lastrow).Specific.Value =  [string]((get-date).AddYears(1).year)

          $oPeriodindicators.Items.item(“1”).click(0)                                                 # create new  期间标识  
          
          
           $SBO_Application.ActivateMenuItem("1596")                                                   # Open Posting Periods
           $oPostingPeriodForm =  $SBO_Application.Forms.ActiveForm
          $oPostingPeriodForm.Items.item(“10000003”).click(0)                                           # click new Periods button

           $newPostingperiodsform = $SBO_Application.Forms.ActiveForm
           $newPostingperiodsform.Items.Item('12').Specific.Value = [string]((Get-Date).AddYears(1).Year)             # new period code
            $newPostingperiodsform.Items.Item('13').Specific.Value = [string]((Get-Date).AddYears(1).Year)              # new period name


            $periodsinds  =  $newPostingperiodsform.Items.Item('39').specific
             $periodsinds.Select((((Get-Date).AddYears(1) ).Year).ToString())                                 # select new year posting periods


               $newPostingperiodsform.Items.Item('14').Specific.Value = [string]((Get-Date).AddYears(1).Year.ToString() + '0101')              # Posting Date from
             
             $newPostingperiodsform.Items.Item('16').Specific.Value = [string]((Get-Date).AddYears(1).Year.ToString()+ '0101')               #  Due Date from
              $newPostingperiodsform.Items.Item('17').Specific.Value = [string]((Get-Date).AddYears(3).Year.ToString()+ '1231')              # due date to 
               $newPostingperiodsform.Items.Item('18').Specific.Value =[string]((Get-Date).Year.ToString()+'0101')                           # doc date from 
                $newPostingperiodsform.Items.Item('19').Specific.Value =[string]((Get-Date).AddYears(1).Year.ToString()+'1231')                           # doc date to 
                $newPostingperiodsform.Items.Item('19').Specific.Value =[string]((Get-Date).AddYears(1).Year.ToString()+'1231')  
               $newPostingperiodsform.Items.Item('36').specific.Value = [string]((Get-Date).AddYears(1).Year.ToString()+ '0101') 
             $newPostingperiodsform.Items.Item('150000051').specific.Value = [string]((Get-Date).AddYears(1).Year) 
                
                $subperiods =  $newPostingperiodsform.Items.Item('22').specific
             $subperiods.Select('M')                                                                            # select months

                     $newPostingperiodsform.Items.item(“1”).click(0)          # click Add
              
          
           # Enable SP contorl in TNMSGS UDO
        Write-Host  -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =""
         fn_SAPB1_SP_control $ticktNum2 'Y'
           
         break  #quit current loop,jump to next site loop
       }
     }
    }
 }

    
  