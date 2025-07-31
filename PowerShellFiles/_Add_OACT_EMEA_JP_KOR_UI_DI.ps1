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


$COA_CSV = IMPORT-CSV \\Client\C$\Temp\COA_China\COA_KOR.csv

function Fn_Add_NewAcct ($sitecode)
{
  
 FOREACH ($r in $COA_CSV){
     $oCoa = $cmp.GetBusinessObject(1)
     # update CoA name
    if ($r.Action.Trim() -EQ "rename account") {
      IF($oCoa.GetByKey($r.'SAP Code')){ 
       
        if ($sitecode -eq 'HG' ){$oCoa.Name =[string]$r.'Account name-English' 
        # $oCoa.ForeignName =[string]$r.'Account name-Local language provided by Local'
         }
        ELSE{ # $oCoa.Name =[string]$r.'Account name-Local language provided by Local'
           $oCoa.ForeignName = [string]$r.'Account name-English'
         }
 
       }
      Write-host  $cmp.CompanyDB " " $r.'SAP Code'  " update Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
     }

      elseif ($r.Action.Trim() -EQ "rename account and controller") {
      IF($oCoa.GetByKey($r.'SAP Code')){ 
       
        if ($sitecode -eq 'HG' ){$oCoa.Name =[string]$r.'Account name-English' 
          if ( $r.Postable -eq 'Y') { $oCoa.ExternalCode = $r.'Controller Code'}
        # $oCoa.ForeignName =[string]$r.'Account name-Local language provided by Local'
         }
        ELSE{ # $oCoa.Name =[string]$r.'Account name-Local language provided by Local'
           $oCoa.ForeignName = [string]$r.'Account name-English'
          if ( $r.Postable -eq 'Y') { $oCoa.ExternalCode = $r.'Controller Code'}
         }
 
       }
      Write-host  $cmp.CompanyDB " " $r.'SAP Code'  " update Acct code and controller code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
     }
      # update default distribution rule
     elseif ($r.Action.Trim() -EQ "change DR Of account") {
      IF($oCoa.GetByKey($r.'SAP Code')){ 
             $oCoa.LoadingFactorCode2 = [string]$r.'Default Location'
             $oCoa.LoadingFactorCode3 = [string]$r.'Default Employee'
            $oCoa.LoadingFactorCode4 = [string]$r.'Default ExpenseType'
       }
      Write-host  $cmp.CompanyDB " " $r.'SAP Code'  " update distribution rule of Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
     }
       
     # Deactivate CoA
     elseif ($r.Action.Trim() -EQ "Inactive"  -and $r.Postable -eq 'Y') {IF($oCoa.GetByKey($r.'SAP Code')){ $oCoa.FrozenFor = 1;$oCoa.ValidFor=0}
        Write-host  $cmp.CompanyDB " "  $r.'SAP Code'  " Inactive Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
       }
     # Change father code
           elseif ($r.Action.Trim() -EQ "Change father code"  ) {IF($oCoa.GetByKey($r.'SAP Code')){
             $oCoa.FatherAccountKey = $r.'SAP father code'}
        Write-host  $cmp.CompanyDB " "  $r.'SAP Code'  " change father Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
       }
     # Add CoA
     elseif ($r.Action.Trim() -EQ "New Account"){      
        $oCoa.code = $r.'SAP Code'
        $oCoa.FatherAccountKey = $r.'SAP father code'.trim()
        if ($sitecode -eq 'HG' ){$oCoa.Name =[string]$r.'Account name-English';   $oCoa.ForeignName =[string]$r.'Account name-Local language provided by Local' }
        ELSE{$oCoa.Name =[string]$r.'Account name-Local language provided by Local';   $oCoa.ForeignName = [string]$r.'Account name-English'}
       
         if ($r.Postable -eq "Y") {  
            $oCoa.ActiveAccount = 1
            $oCoa.ExternalCode = $r.'Controller Code'
             $oCoa.AcctCurrency = $r.Currency
        if ($r.'Controlled CoA' -eq 'Y') {$oCoa.BlockManualPosting = 1 } else {$oCoa.BlockManualPosting = 0} 
        if ($r.'BP Control Account' -eq 'Y') {$oCoa.LockManualTransaction = 1}
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
              if(![string]::IsNullOrEmpty($r.'U_VES_AcctRecon')) {$oCoa.UserFields.Fields.Item("U_VES_AcctRecon").Value = $r.'U_VES_AcctRecon'}
            
    
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

Fn_Add_NewAcct "HG"