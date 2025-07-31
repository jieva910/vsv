

    $cmp = New-Object -COMobject "SAPbobsCOM.Company"


function UI_DI_Conn
{
 
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
    
             $sCookie = $cmp.GetContextCookie()
             $sConnectionContext = $SBO_Application.Company.GetConnectionContext($sCookie)
             If ($cmp.Connected ){$cmp.Disconnect()}
             return $cmp.SetSboLoginContext($sConnectionContext)
        }

    # Connect to SBO via DI API

    Function ConnectTcmp {
       Return $cmp.Connect()
    }

    # connect to DI 
   
        if (SetConnectionContext -ne 0 ) {$SBO_Application.MessageBox("Failed setting a connection to DI API");break}
        if (ConnectTcmp -ne 0 ) {$SBO_Application.MessageBox("Failed connecting to the company's Data Base") ; break}
    Write-Host -ForegroundColor Cyan "DI Connected To: " $cmp.LicenseServer $cmp.CompanyName
}

UI_DI_Conn

function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}

# Control the SP of blocking Super user
Function fn_SAPB1_SP_control($ticknum,$YesNO,$site)
  {   
   $oCompServic = $cmp.GetCompanyService()
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

           Release-Ref ($oCompServic)
      }
}




function Fn_Add_NewAcct ($sitecode)
{
  
   
 $COA_CSV =Import-Csv \\dg-fs-usr02\KKUSERS\jieva\temp\COAKB.csv

 FOREACH ($r in $COA_CSV){
     $oCoa = $cmp.GetBusinessObject(1)
     # update CoA
    if ($r.Action.Trim() -EQ "rename account") {
      IF($oCoa.GetByKey($r.'SAP Code')){ 
       
        if ($sitecode -eq 'HG' ){$oCoa.Name =[string]$r.'Account name-English' 
        $oCoa.ForeignName =[string]$r.'Account name-Local language provided by Local'
         }
        ELSE{$oCoa.Name =[string]$r.'Account name-Local language provided by Local'
           $oCoa.ForeignName = [string]$r.'Account name-English'
         }
 
       }
      Write-host  $cmp.CompanyDB " " $r.'SAP Code'  " update Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
     }
     # Deactivate CoA
     elseif ($r.Action.Trim() -EQ "Inactive"  -and $r.Postable -eq 'Y') {IF($oCoa.GetByKey($r.'SAP Code')){ $oCoa.FrozenFor = 1;$oCoa.ValidFor=0}
        Write-host  $cmp.CompanyDB " "  $r.'SAP Code'  " Inactive Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
       }
     # Add CoA
     elseif ($r.Action.Trim() -EQ "New Account"){      
        $oCoa.code = $r.'SAP Code'
        $oCoa.FatherAccountKey = $r.'SAP father code'
        if ($sitecode -eq 'HG' ){$oCoa.Name =[string]$r.'Account name-English';   $oCoa.ForeignName =[string]$r.'Account name-Local language provided by Local' }
        ELSE{$oCoa.Name =[string]$r.'Account name-Local language provided by Local';   $oCoa.ForeignName = [string]$r.'Account name-English'}
        $oCoa.AcctCurrency = $r.Currency
        if ($r.'Controlled CoA' -eq 'Y') {$oCoa.BlockManualPosting = 1 } else {$oCoa.BlockManualPosting = 0} 

         if ($r.Postable -eq "Y") {  
            $oCoa.ActiveAccount = 1
            $oCoa.ExternalCode = $r.'Controller Code'
            SWITCH ($r.AcctType)
            {
              "I" {$oCoa.AccountType = 0
                     $oCoa.DistributionRuleRelevant = 1
                    $oCoa.DistributionRule2Relevant = 1
                    $oCoa.DistributionRule3Relevant = 1
                    $oCoa.DistributionRule4Relevant = 1
                    $oCoa.DistributionRule5Relevant = 1
                  }   #  at_Revenues = 0 
              "E" {$oCoa.AccountType = 1
                    $oCoa.DistributionRuleRelevant = 1
                    $oCoa.DistributionRule2Relevant = 1
                    $oCoa.DistributionRule3Relevant = 1
                    $oCoa.DistributionRule4Relevant = 1
                    $oCoa.DistributionRule5Relevant = 1
                    $oCoa.LoadingFactorCode4 = $r.'Default ExpenseType'}   # at_Expenses =1 
              Default {$oCoa.AccountType = 2} # at_Other =2
            }

            if(![string]::IsNullOrEmpty($r.'Allocation of SS')) {$oCoa.UserFields.Fields.Item("u_ves_alloc_ss").Value = $r.'Allocation of SS'}
    
             $oRs = $cmp.GetBusinessObject(300)
            $oRs.DoQuery("SELECT AbsId FROM OACG  WHERE Name ='$($r.'Account Category')'")
            If (!$oRs.EOF){ $oCoa.Category = $oRs.Fields.Item(0).Value}
             
         }
         ELSE {$oCoa.ActiveAccount = 0}
     
         Write-Output ($cmp.CompanyDB +" "+  $r.'SAP Code' + " Added with err code " + $oCoa.Add() + " " + $cmp.GetLastErrorDescription())
        
        }
     Release-Ref $oCoa
}
   
   
}

Fn_Add_NewAcct 'os'