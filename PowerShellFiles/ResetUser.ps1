
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1
$cmp = New-Object -ComObject "sapbobscom.company"


$SourceSite    = "sq"
$usercode = 'zhangann'
$ticknum       = 'INC0222801'

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp $SourceSite

 
# Disable SP control 
fn_SAPB1_SP_control $ticknum 'N'  $SourceSite 



 $usr = $cmp.GetBusinessObject('12')
    $rs = $cmp.GetBusinessObject('300') #recordset

    $rs.doquery(“select userid from ousr where user_code='" +$usercode + "'")
    if ($rs.EoF -eq $false)
     {  $uid = $rs.Fields.Item(0).value
        if ($usr.getbykey($uid) -eq $true )
        {$usr.locked =0 ;
           # $usr.UserPassword ='Ves-123456'
          #$usr.SUPERUSER=1
            $err_msg = $usr.Update()
           if ($err_msg -eq 0) 
            { Write-Host $usercode  Unlock OK}
           else{Write-Host $cmp.GetLastErrorDescription()}
         }
     }
      else{Write-Host $usercode  not exiting}

$ticknum =''
fn_SAPB1_SP_control $ticknum 'Y'  $SourceSite 

