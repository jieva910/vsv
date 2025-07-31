
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
#---------------enable InitialSessionState-----
$InitialSessionState=[System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
#-------------get all function in current session---
 Get-ChildItem Function:\ | Where-Object {$_.Name -notlike "*:*"}|select name -ExpandProperty name|
 ForEach-Object{$definition=Get-Content "function:\$_" -ErrorAction Stop
 $function=New-object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $_ ,$definition
 $InitialSessionState.Commands.Add($function) }

#---------------enable runspacepool-----
$Runspacepool=[runspacefactory]::createrunspacepool($InitialSessionState)
[void]$Runspacepool.SetMinRunspaces(1)
[void]$Runspacepool.SetMaxRunspaces(10)
$Runspacepool.Open()
$powershell=[powershell]::create()
$powershell.RunspacePool=$Runspacepool
$hash=[hashtable]::Synchronized(@{})

$jobs = New-Object System.Collections.ArrayList
$starttime = Get-DATE

function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}

function fn_UpdateBP {
    param ($cmp ,$bpcode)
     $row_data=$bpcode -split ","
     $oBP=$cmp.getbusinessobject(2)
    if ( $oBP.GetByKey($row_data[0].Trim()) ) {
        $oBP.CreditLimit =$row_data[1]
        $oBP.MaxCommitment = $row_data[2]
                $outputcode = $oBP.Update()
        $outlog =$cmp.CompanyDB + ' '+ $bpcode + ' updated  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
        return $outlog
      }
}
#----OPEN UDF CSV FILE-------------
 $CSVfile = Get-Content -ReadCount 0 -LiteralPath C:\Temp\sz_BPCODE.csv

 $SAP_SiteConnS = @{SZ =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1-SZ";sapuser="jieva";pwd="Ves-1234"}
#TK =@{ Lic="DG-SAPLIC01";db="DG-APAC-SAP01";dbtype="8";cmp="SAPB1_TK";sapuser="jieva";pwd="Ves1234"}
}
  
  #Connect to SAPB1 in loop using above site 
 $SAP_SiteConnS.Keys | Sort-Object | ForEach-Object { 
        $cmpServer = $SAP_SiteConnS[$_]['db']
        $cmpCompanyDB = $SAP_SiteConnS[$_]['cmp']
        $cmpDbServerType = $SAP_SiteConnS[$_]['dbtype']
        $cmpUserName = $SAP_SiteConnS[$_]['sapuser']
        $cmpPassword =$SAP_SiteConnS[$_]['pwd']
        $cmpLicenseServer = $SAP_SiteConns[$_]['Lic']
     
  
         $cmp.Server = $cmpServer
        $cmp.CompanyDB =$cmpCompanyDB
        $cmp.DbServerType = $cmpDbServerType
        $cmp.UserName = $cmpUserName
        $cmp.Password =$cmpPassword
    
        if ($_ -eq 'PG' -OR $_ -eq 'KH')
          {$cmp.DbUserName="Butterfly" ;$cmp.DbPassword="buTterF1y"}
        elseif ($_ -eq'RK') {$cmp.DbUserName="BoomRang";$cmp.DbPassword="B00mrang"}
        else  {$cmp.UseTrusted = $True}
        $cmp.LicenseServer = $cmpLicenseServer

        [void]$cmp.Connect()
        if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() 
              EXIT} }
     #loop the csv file and update them with multiple threading
     ForEach ($row in $CSVfile){
              $powershell=[powershell]::Create()
              $powershell.RunspacePool=$RunSpacePool
              
           $paramlist=@{
                 com=$cmp
                 Param2=$row 
                 }
          $myscript= {
            Param( $com, 
                 $Param2,
                 $hash
              )
            $cell=$Param2 -split ","
            $outlog2=fn_UpdateBP $com $Param2 
            $hash[$cell[0]]=$outlog2 
         }

        $powershell.AddScript($myscript).addargument($hash)|Out-Null
       $powershell.AddParameters($paramlist) |Out-Null
       $handle = $powershell.BeginInvoke()
       $tmp = '' | select powershell,handle
       $tmp.powershell=$powershell
       $tmp.handle=$handle
       $jobs.Add($tmp) | Out-Null

      }
    
    $return=$jobs|ForEach-Object{
              $_.powershell.endinvoke($_.handle)
              $_.powershell.dispose()
     }
    $jobs.Clear() | Out-Null    #release Memory which stored the powershell's handle
    $return                    #get the return value ,that is the value stored in Hash table
    $cmp.Disconnect() | Out-Null   #disconnected from current SAPB1 
    $hash.keys|Sort-Object | ForEach-Object{ $hash[$_] } | Out-File -Append 'c:\temp\updated_sz_Creditlimit.log'

    $endtime =Get-Date

     Write-Host -ForegroundColor Red ('running time is :' +($endtime-$starttime).totalseconds)


 Release-Ref ($cmp)
