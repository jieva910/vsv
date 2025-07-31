
cls

$site = 'SZTST'

. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

Fn_ConnectSAPB1 $site

fn_SAPB1_SP_control 'test' 'N' 'SZ'



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
[void]$Runspacepool.SetMaxRunspaces(5)
$Runspacepool.Open()
$powershell=[powershell]::create()
$powershell.RunspacePool=$Runspacepool
$hash=[hashtable]::Synchronized(@{})

$jobs = New-Object System.Collections.ArrayList


function fn_UpdateBP {
    param ($cmp,$CardCode,$creditline,$DebtLine)
     
     $oBP=$cmp.getbusinessobject(2)
    if ( $oBP.GetByKey($CardCode) ) {
        $oBP.CreditLimit =$creditline
        $oBP.MaxCommitment = $DebtLine
        $outputcode = $oBP.Update()
        $outlog =$cmp.CompanyDB + ' '+ $CardCode + ' updated  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
        return $outlog
      }
}

$starttime = Get-DATE


# $CSVfile = Get-Content -ReadCount 0 -LiteralPath C:\Temp\cK_BPCODE.csv


 $CSVfile = Import-Csv -LiteralPath C:\Temp\SZBP.CSV -Delimiter ','
 
 # ForEach ($row in $CSVfile){fn_UpdateBP $row.CardCode $row.CreditLine $row.DebtLine }


#loop the csv file and update them with multiple threading
ForEach ($row in $CSVfile){
              $powershell=[powershell]::Create()
              $powershell.RunspacePool=$RunSpacePool
              
           $paramlist=@{
                 oComp=$cmp
                 CardCode=$row.cardcode
                 creditline = $row.CreditLine
                  DebtLine = $row.DebtLine

                 }
          $myscript= {
            Param( $oComp, 
                 $CardCode,
                 $creditline,
                 $DebtLine,
                 $hash
              )
            $hash_Key=$CardCode
            $outlog2=fn_UpdateBP $oComp  $CardCode $creditline  $DebtLine
            $hash[$hash_Key]=$outlog2 
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
  
    $endtime =Get-Date

    Write-Host -ForegroundColor Red ('running time is :' +($endtime-$starttime).totalseconds)
   
    $log_file=  $hash.keys| ForEach-Object{ $hash[$_] } 
  
   Out-File -FilePath 'C:\TEMP\update.LOG' -Append -Encoding utf8 -InputObject $log_file

   $cmp.Disconnect() | Out-Null   #disconnected from current SAPB1 
    