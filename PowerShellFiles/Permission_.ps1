
cls

$SourceSite    = "hg"

$ticknum       = 'super user audit'
# load sapb1 di connection lib


.  C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite

 
# Disable SP control 
fn_SAPB1_SP_control $ticknum 'N'  $SourceSite 





$PermissionIDs = "1,13,14,31,36,37,38,39,40,41,142,271,400,442,472,610,920,1101,UP_MAIN_HEADER"
$oBobi         =$cmp.GetBusinessObject(305) # SBObob object 
$usr           = $cmp.GetBusinessObject('12')
$rs            = $cmp.GetBusinessObject('300') #recordset

$usrcodes       = "xiapan,yakerlyn,ayresrob,borgojoh,godinfre,vandeluc,nairaji,tajandee,stenpdir,jieva2,BEOSProduWouweFil,panxia"


Foreach ($usrcode in $usrcodes -split ",")
{
   # Remove super role 
      $rs.doquery(“select userid from ousr where SUPERUSER ='Y' and  user_code='" +$usrcode + "'")
    if ($rs.EoF -eq $false)
     {  $uid = $rs.Fields.Item(0).value
        if ($usr.getbykey($uid) -eq $true )
        {$usr.locked =0 ;
         $usr.SUPERUSER =0 
       Write-Host   "$usrcode Super Role Removed with Errorcode:" $usr.Update()  $cmp.GetLastErrorDescription()
         }
     
       # Grant User Full authorizations
        foreach($ID in $PermissionIDs -split ",")
        {
          $oBobi.SetSystemPermission($usrcode, $ID, 1); # 38 is Business parter module, 1 is full ,2 is read-only
          $cmp.GetLastErrorDescription()
        }
     }
      else{Write-Host $usrcode  not exiting}
    
   
}



# $mUserPermission = $cmp.getbusinessobject(214)   #  UserPermissionTree object 


 
# Enable SP control 
$ticknum = "   "
fn_SAPB1_SP_control $ticknum 'Y' $SourceSite 


