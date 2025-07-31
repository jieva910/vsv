<#
  Purpose :  update OACT foreign name within china sites
  Date    : 2021/1
#>

cls
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'

$ticktNum = "newCOA"



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
  
   
 $COA_CSV =Import-Csv C:\Temp\COA_China\COACN2408.csv

 FOREACH ($r in $COA_CSV){
     $oCoa = $cmp.GetBusinessObject(1)
     # update CoA name
    if ($r.Action.Trim() -EQ "rename account") {
      IF($oCoa.GetByKey($r.'SAP Code')){ 
       
        if ($sitecode -eq 'HG' ){$oCoa.Name =[string]$r.'Account name-English' 
      #  $oCoa.ForeignName =[string]$r.'Account name-Local language provided by Local' 
         }
        ELSE{ # $oCoa.Name =[string]$r.'Account name-Local language provided by Local'
           $oCoa.ForeignName = [string]$r.'Account name-English'
         }
 
       }
      Write-host  $cmp.CompanyDB " " $r.'SAP Code'  " update Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
     }

      elseif ($r.Action.Trim() -EQ "rename account and controller") {
      IF($oCoa.GetByKey($r.'SAP Code')){ 
       
       # if ($sitecode -eq 'HG' ){$oCoa.Name =[string]$r.'Account name-English' 
          if ( $r.Postable -eq 'Y') { $oCoa.ExternalCode = [string]$r.'Controller Code'}
        # $oCoa.ForeignName =[string]$r.'Account name-Local language provided by Local'
        # }
       # ELSE{ # $oCoa.Name =[string]$r.'Account name-Local language provided by Local'
         #  $oCoa.ForeignName = [string]$r.'Account name-English'
         # if ( $r.Postable -eq 'Y') { $oCoa.ExternalCode = $r.'Controller Code'}
         #}
 
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
     elseif ($r.Action.Trim() -EQ "inactive"  -and $r.Postable -eq 'Y') {IF($oCoa.GetByKey($r.'SAP Code')){ $oCoa.FrozenFor = 1;$oCoa.ValidFor=0 ;$ocoa.FrozenRemarks=$ticktNum}
        Write-host  $cmp.CompanyDB " "  $r.'SAP Code'  " Inactive Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
       }
     # Change father code
           elseif ($r.Action.Trim() -EQ "Change father code"  ) {IF($oCoa.GetByKey($r.'SAP Code')){
             $oCoa.FatherAccountKey = $r.'SAP father code'}
        Write-host  $cmp.CompanyDB " "  $r.'SAP Code'  " change father Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
       }
        # Change external code
           elseif ($r.Action.Trim() -EQ "Change Controller code"  ) {IF($oCoa.GetByKey($r.'SAP Code')){
             $oCoa.ExternalCode= $r.'Controller Code'
              Write-host  $cmp.CompanyDB " "  $r.'SAP Code'  " change external code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
             }
           }
     # Add CoA
     elseif ($r.Action.Trim() -EQ "New Account"){      
        $oCoa.code = $r.'SAP Code'
        $oCoa.FatherAccountKey = $r.'SAP father code'.trim()
        if ($sitecode -eq 'HG' ){$oCoa.Name =[string]$r.'Account name-English';   $oCoa.ForeignName =[string]$r.'Account name-Local language provided by Local' }
        ELSE{$oCoa.Name =[string]$r.'Account name-Local language provided by Local';   $oCoa.ForeignName = [string]$r.'Account name-English'}
       
         if ($r.Postable -eq "Y"  ) {  
            $oCoa.ActiveAccount = 1
            $oCoa.ExternalCode = $r.'Controller Code'
             $oCoa.AcctCurrency = $r.Currency
        if ($r.'Controlled CoA' -eq 'Y') {$oCoa.BlockManualPosting = 1 } else {$oCoa.BlockManualPosting = 0} 
        if ($r.'BP Control Account' -eq 'Y') {$oCoa.LockManualTransaction = 1}
        if ( $r.'Bank / cash' -eq 'Y') {$oCoa.CashAccount = 1}
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
                    $oCoa.LoadingFactorCode4 = $r.'Default ExpenseType'
                    $oCoa.LoadingFactorCode2 = [string]$r.'Default Location'
             $oCoa.LoadingFactorCode3 = [string]$r.'Default Employee'
                    }   # at_Expenses =1 

              Default {$oCoa.AccountType = 2} # at_Other =2
            }

            if(![string]::IsNullOrEmpty($r.'Allocation of SS')) {$oCoa.UserFields.Fields.Item("u_ves_alloc_ss").Value = [string]$r.'Allocation of SS'}
              if(![string]::IsNullOrEmpty($r.'U_VES_AcctRecon')) {$oCoa.UserFields.Fields.Item("U_VES_AcctRecon").Value = [string]$r.'U_VES_AcctRecon'}
            
    
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

function fn_AddCOAXML
{
       Write-Host "Start add new account code  ......"
  $xml_path = @('C:\Temp\999996-40.xml','C:\Temp\999998-22.xml')

  foreach($p in $xml_path)
  {
     $oCoa = $cmp.GetBusinessObjectFromXML($p,0)
     Write-Output ($cmp.CompanyDB +" Added with err code " + $oCoa.Add() + " " + $cmp.GetLastErrorDescription())
     Release-Ref $oCoa
  }

   
}

$SAP_SiteConnS = @{ 
 AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-123456";DbUserName="Butterfly";DbPassword="buTterF1y"}
 BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
YK =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_YK";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
 CS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_CS";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
KT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_KT";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
 WN =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WN";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
 HG =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_HG";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
# WE =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WE";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
# SQ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_SQ";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}  
WG =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WG";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
WV =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WV";sapuser="\";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"}
#  CSTST =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPstg91";dbtype="8";cmp="SAPB1_CS_tst";sapuser="b1i";pwd="Ves-123456";DbUserName="Butterfly";DbPassword="buTterF1y"}
# SZTST =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPstg91";dbtype="8";cmp="SAPB1_SZ_tst";sapuser="b1i";pwd="Ves-123456";DbUserName="Butterfly";DbPassword="buTterF1y"}
# CSTST2 =@{ Lic="SZ-TSTSAPLIC92";db="SZ-SAPtst82";dbtype="8";cmp="SAPB1_CS_tst2";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}

 # CK =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_CK";sapuser="CORP\jieva";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"} 
#   TK =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_TK";sapuser="CORP\jieva";pwd="";DbUserName="Butterfly";DbPassword="buTterF1y"} 
 #  SL =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_SL";sapuser="jieva";pwd="Ves-12345";DbUserName="Butterfly";DbPassword="buTterF1y"}  
 #  KB =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_KB";sapuser="jieva";pwd="Ves-12345";DbUserName="Butterfly";DbPassword="buTterF1y"}   
   
}


 $SAP_SiteConnS.Keys|Sort-Object|  ForEach-Object { 
        $cmpServer = $SAP_SiteConnS[$_]['db']
        $cmpCompanyDB = $SAP_SiteConnS[$_]['cmp']
        $cmpDbServerType = $SAP_SiteConnS[$_]['dbtype']
        $cmpUserName = $SAP_SiteConnS[$_]['sapuser']
        $cmpPassword =$SAP_SiteConnS[$_]['pwd']
        $cmpLicenseServer = $SAP_SiteConns[$_]['Lic']
        $cmpdbuser=$SAP_SiteConns[$_]['DbUserName']
        $cmpdbpwd=$SAP_SiteConns[$_]['DbPassword']
  
         $cmp.Server = $cmpServer
        $cmp.CompanyDB =$cmpCompanyDB
        $cmp.DbServerType = $cmpDbServerType
        $cmp.UserName = $cmpUserName
        $cmp.Password =$cmpPassword
       # $cmp.DbUserName=$cmpdbuser
       # $cmp.DbPassword=$cmpdbpwd
        $cmp.UseTrusted=$true
        $cmp.LicenseServer = $cmpLicenseServer

        [void]$cmp.Connect()
       
        if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;Continue
              } 
        else { Write-Host -ForegroundColor Cyan $cmp.CompanyDB connected successfully}

         #disable SAPB1 TN sp control
         Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"

         fn_SAPB1_SP_control $ticktNum 'N' $_

        

          # fn_Update_AcctName
         Fn_Add_NewAcct $_

         


         #Enable sapb1 TN SP control
         Write-Host -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =''
         fn_SAPB1_SP_control $ticktNum2 'Y' $_
        Write-Host "End processing ."       
       
       $cmp.Disconnect()

  } 

#Release-Ref($cmp)  