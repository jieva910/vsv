
<#
  purpose : 使用2个或多个SAPB1 账号 连接 COM object , 然后同时运行这几个实列。
  Date    : 2021.01
  update  : 2021.05.12  测试150行数据，性能反而不好。
#>


$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
$cmp2 = New-Object -COMObject 'SAPbobsCOM.Company'

$cmp.Server = "SZ-SAPtst82"
$cmp.CompanyDB ="SAPB1_SZ_TST"
$cmp.DbServerType = 8
$cmp.UserName = "/"
$cmp.Password ="Ves-1234"
# $cmp.DbUserName=$cmpdbuser
# $cmp.DbPassword=$cmpdbpwd
$cmp.UseTrusted=$true
$cmp.LicenseServer = "SZ-TSTSAPLIC92"

[void]$cmp.Connect()
       
if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;break
              } 
else { Write-Host -ForegroundColor Cyan $cmp.CompanyDB connected successfully}



$cmp2.Server = "SZ-SAPTST82"
$cmp2.CompanyDB ="SAPB1_SZ_TST"
$cmp2.DbServerType = 8
$cmp2.UserName = "shenste"
$cmp2.Password ="Ves-123456"
# $cmp2.DbUserName=$cmp2dbuser
# $cmp2.DbPassword=$cmp2dbpwd
$cmp2.UseTrusted=$true
$cmp2.LicenseServer = "SZ-TSTSAPLIC92"

[void]$cmp2.Connect()
       
if(-not $cmp2.Connected) {Write-Host $cmp2.GetLastErrorDescription() ;break
              } 
else { Write-Host -ForegroundColor Cyan $cmp2.username connected successfully}


$toexecute1 = {
  Param($cmp)
    $CSV = Import-Csv 'C:\Temp\SZBP.csv'
    foreach ($x in $csv){
         $oBP=$cmp.getbusinessobject(2)
          $cad=$x.CardCode
         $t1=$x.CreditLine
         $t2= $x.DebtLine
        if ( $oBP.GetByKey($cad) ) {
            $oBP.CreditLimit = $t1
            $oBP.MaxCommitment =$t2
            $outputcode = $oBP.Update()
           [PSCustomobject] @{ 
                Success = $outputcode -eq 0       # 假设传入参数5时失败，其余成功
                Data = "结果 $outputcode"     # 假设Data是执行结果，带上传入参数以区分
                ThreadId = [AppDomain]::GetCurrentThreadId()   # 当前线程ID  
            }
          }
    }
}

 $toexecute2 =    {  Param ($cmp2)
    $CSV2 = Import-Csv 'C:\Temp\SZBP2.csv'
    foreach ($x2 in $CSV2){
         $oBP=$cmp2.getbusinessobject(2)
           $cad=$x2.CardCode
         $t1=$x2.CreditLine
         $t2= $x2.DebtLine
        if ( $oBP.GetByKey($cad) ) {
            $oBP.CreditLimit = $t1
            $oBP.MaxCommitment =$t2
            $outputcode = $oBP.Update()
           [PSCustomobject] @{ 
                Success = $outputcode -eq 0       # 假设传入参数5时失败，其余成功
                Data = "结果 $outputcode"     # 假设Data是执行结果，带上传入参数以区分
                ThreadId = [AppDomain]::GetCurrentThreadId()   # 当前线程ID  
            }
          }
    }
}

$starttime = Get-Date

$a =  [PowerShell]::Create().AddScript($toexecute1).AddArgument($cmp)
$b =  [PowerShell]::Create().AddScript($toexecute2).AddArgument($cmp2)
#$c =  [PowerShell]::Create().AddScript{sleep 5;'c done'}
$r1,$r2 = ($a,$b).begininvoke() # run in background
$a.EndInvoke($r1); $b.EndInvoke($r2);  # wait
($a,$b).streams.error # check for errors
($a,$b).dispose() # clean


$endtime =Get-Date

Write-output ("Running Time is :" + ($endtime-$starttime).TotalSeconds)