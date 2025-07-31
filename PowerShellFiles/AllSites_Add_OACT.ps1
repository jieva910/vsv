<#
  Purpose :  update OACT foreign name within china sites
  Date    : 2021/1
#>

cls
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'

$ticktNum = "inactive coa per julia jiang requst"

$AcctCode = "565300-01"
$AcctForeignName = "Loss on Write Off of F.A.> GBP250K"

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


Function fn_Update_AcctName 
{
    $oAcct = $cmp.GetBusinessObject(1) # oChartOfAccounts = 1
    
$COA_CSV =Import-Csv C:\Temp\COA.csv
    FOREACH ($r in $COA_CSV){
     IF($oAcct.GetByKey($r.'SAP Code'))
     {
       # $oAcct.ForeignName = $AcctForeignName

       $oAcct.LockManualTransaction = 1 # set as Control account
       Write-host  $cmp.CompanyDB " update Acct code with error code : " $oAcct.Update() "and error description: "$cmp.GetLastErrorDescription() 
     }
    }
}

function Fn_Add_NewAcct
{
  
     Write-Host "Start add new account code  ......"
$COA_CSV =Import-Csv C:\Temp\COA-inactive.csv

 FOREACH ($r in $COA_CSV){
     $oCoa = $cmp.GetBusinessObject(1)
    if ($r.Action.Trim() -EQ "rename account") {
      IF($oCoa.GetByKey($r.'SAP Code')){ 
       if ($r.Postable -eq 'N'){$oCoa.ForeignName = $r.'Account name-English';$oCoa.Name =$r.'Account name-English'
          }
        elseif ($r.Postable -eq 'Y'){$oCoa.ForeignName = $r.'Account name-English';$oCoa.Name =$r.'Account name-English'
           if ($r.'Controlled CoA' -eq 'Y'){ $oCoa.BlockManualPosting = 1 }
           $oCoa.UserFields.Fields.Item("u_ves_alloc_ss").Value = $r.u_ves_alloc_ss         
                }
       }
          Write-host  $cmp.CompanyDB " " $r.'SAP Code'  " update Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
     }
     elseif ($r.Action.Trim() -EQ "Inactive"  -and $r.Postable -eq 'Y') {IF($oCoa.GetByKey($r.'SAP Code')){ $oCoa.FrozenFor = 1;$oCoa.ValidFor=0}
        Write-host  $cmp.CompanyDB " "  $r.'SAP Code'  " Inactive Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
       }
   
     if ($r.Action.Trim() -EQ "New Account"){      
        $oCoa.code = $r.'SAP Code'
        $oCoa.FatherAccountKey = $r.'SAP father code'
        $oCoa.Name =  $r.'Account name-Local language'
        $oCoa.ForeignName = $r.'Account name-English'
        $oCoa.AcctCurrency = $r.Currency
        
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
                    $oCoa.DistributionRule5Relevant = 1}   # at_Expenses =1 
              Default {$oCoa.AccountType = 2} # at_Other =2
            }

           # $oCoa.UserFields.Fields.Item("u_ves_alloc_ss").Value = $r.u_ves_alloc_ss
    
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



$SAP_SiteConnS = @{ 
 AS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_AS";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="CORP\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
 CS =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_CS";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
KT =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_KT";sapuser="CORP\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="jieva";pwd="Ves-1234";DbUserName="Butterfly";DbPassword="buTterF1y"}
#WE =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WE";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
SQ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_SQ";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}  
WN =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_WN";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
WG =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WG";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
WV =@{ Lic="SZ-SAPLIC92";db="WG-SAP01";dbtype="8";cmp="SAPB1_WV";sapuser="jieva";pwd="vesint99";DbUserName="Butterfly";DbPassword="buTterF1y"}
 HG =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_HG";sapuser="corp\jieva";pwd="Vesint-999";DbUserName="Butterfly";DbPassword="buTterF1y"}
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
         Fn_Add_NewAcct

         


         #Enable sapb1 TN SP control
         Write-Host -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
        $ticktNum2 =''
         fn_SAPB1_SP_control $ticktNum2 'Y' $_
        Write-Host "End processing ."       
       
        $cmp.Disconnect()

  } 

#Release-Ref($cmp)  