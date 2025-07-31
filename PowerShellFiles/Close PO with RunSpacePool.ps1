
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
#---------------enable InitialSessionState-----
$InitialSessionState=[System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
#-------------get all function in current session---
 Get-ChildItem Function:\ | Where-Object {$_.Name -notlike "*:*"}|select name -ExpandProperty name|
 ForEach-Object{$definition=Get-Content "function:\$_" -ErrorAction Stop
 $function=New-object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $_ ,$definition
 $InitialSessionState.Commands.Add($function) }

#---------------enable Runspacepool-----
$Runspacepool=[runspacefactory]::createRunspacepool($InitialSessionState)
[void]$Runspacepool.SetMinRunspaces(1)
[void]$Runspacepool.SetMaxRunspaces(10)
$Runspacepool.Open()
$powershell=[powershell]::create()
$powershell.Runspacepool=$Runspacepool
$hash=[hashtable]::Synchronized(@{})

$jobs = New-Object System.Collections.ArrayList
$starttime = Get-DATE

function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}

function fn_ClosePo {
    param ($cmp ,$ponumber)
    $po = $cmp.GetBusinessObject(22)
    $rs = $cmp.GetBusinessObject('300') #recordset

    $rs.doquery(“select docentry from opor where docnum=" +$ponumber + "")
    if ($rs.EoF -eq $false)
     {  $poDocentry = $rs.Fields.Item(0).value
        if ($po.getbykey($poDocentry) -eq $true )
        {
          $outputcode = $po.CLOSE()
        $outlog =$cmp.CompanyDB + ' '+ $ponumber.tostring() + ' updated  with error code:' +   $outputcode.tostring() + ' and error description is:' + $cmp.GetLastErrorDescription()
         return $outlog
         }
      }
}
#----OPEN UDF CSV FILE-------------
 $CSVfile = Get-Content -ReadCount 0 -LiteralPath C:\Temp\BY_PO.csv

 $SAP_SiteConnS = @{BY =@{ Lic="SZ-SAPLIC92";db="SZ-SAP01";dbtype="8";cmp="SAPB1_BY";sapuser="jieva";pwd="Ves-1234"}
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
              $powershell.Runspacepool=$Runspacepool
              
           $paramlist=@{
                 com=$cmp
                 Param2=$row 
                 }
          $myscript= {
            Param( $com, 
                 $Param2,
                 $hash
              )
           $outlog2=fn_ClosePo $com $Param2 
           $hash[$Param2]=$outlog2 
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
  
   $outTime=Get-Date -format 'yyyyMMddHms'
 $outpath = "c:\temp\$outTime.log"
 $log= $hash.Keys | Sort-Object | ForEach-Object{ $hash[$_] }   
  set-Content $outpath -Value $log -Encoding Unicode

    $endtime =Get-Date

     Write-Host -ForegroundColor Red ('running time is :' +($endtime-$starttime).totalseconds)

[void]$cmp.Disconnect()
 Release-Ref ($cmp)
