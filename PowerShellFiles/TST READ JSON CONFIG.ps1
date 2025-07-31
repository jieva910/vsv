$json = Get-Content $PSScriptRoot\config.json |ConvertFrom-Json
# $json.DbServerType
$ServerList = @{
"SZ-SAPLIC92" =@{db="SLDModel.SLDData";LicenseSvr = "Y"}
"SZ-SAP01" = @{db="SAPB1_WN";dbtype=$json.DbServerType;sapid="Montova"; sappwd ="ButterWN";lic="SZ-SAPLIC92"}
"WG-SAP01" = @{db="SAPB1_WG";dbtype=$json.DbServerType;sapid="Montova"; sappwd ="ButterWG";lic="SZ-SAPLIC92"}
}

$ServerList.Keys  | ForEach-Object  {    
    # initial arguments
  
  $ServerList[$_]['db']
    $ServerList[$_]['LicenseSvr']
 $ServerList[$_]['dbtype']
      }