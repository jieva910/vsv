cls

$cmp = New-Object -ComObject "sapbobscom.company"

$ticknum = 'INC0383147'

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$csv = Import-Csv C:\Dell\cnCOAReview.csv -Delimiter ","
$sites = "AS","BY","YK"



foreach($site in $sites -split ",")
{
  # SAPB1 DI connect to specific site
  Fn_ConnectSAPB1 $cmp $site

  # Disable SP control 
  fn_SAPB1_SP_control $ticknum 'N'  $site 

  $oCoa = $cmp.GetBusinessObject(1)
    
  foreach($row in $csv  | Where-Object{$_.site -eq $site} )
  {
   
           # Deactivate CoA
         if ($row.Action -EQ "inactive") {IF($oCoa.GetByKey($row.'SAP Code')){ If($oCoa.FrozenFor -ne 1){ $oCoa.FrozenFor = 1;$oCoa.ValidFor=0}}
        Write-host  $cmp.CompanyDB " "  $row.'SAP Code'  " Inactive Acct code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
       }
   

     <#  if ($row.Action -EQ "rename account and controller") {
      IF($oCoa.GetByKey($row.'SAP Code')){ 
          $oCoa.ExternalCode = [string]$row.'Controller Code'
          Write-host  $cmp.CompanyDB " " $row.'SAP Code'.Trim()  " update  controller code with error code : " $oCoa.Update() "and error description: "$cmp.GetLastErrorDescription() 
            }
	    } #>
 
  }
 
  # Enable SP control
  $ticknum2='' 
 fn_SAPB1_SP_control $ticknum2 'Y'  $site 

  $CMP.Disconnect()
  Start-Sleep 5
}