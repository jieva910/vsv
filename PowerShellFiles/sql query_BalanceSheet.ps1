<#
  Purpose: connect to SQL Server with Paralle and run the sql 
  date:20200602
#>
 $filepath ="C:\shared\PShell\pshell_sql\DBServer_SAPDB.txt"
 $ComPList = Get-Content -ReadCount 0  "$filepath" 
 $RowCounts=$ComPList.Count-1 # start with 0 in Array

 $SQL="
WITH BSQuery AS (

SELECT T2.[GroupMask],T1.[Account], T2.[AcctName], T2.Levels

, Sum(T1.[Debit]-T1.[Credit]) [CurrentPeriod]

FROM OJDT T0

INNER JOIN JDT1 T1 ON T0.[TransId] = T1.[TransId]

INNER JOIN OACT T2 ON T1.[Account] = T2.[AcctCode]

WHERE T2.[GroupMask] in (1,2,3)

and T0.[RefDate] <= convert(VARCHAR(10),getdate(),112)

GROUP BY T2.[GroupMask], T1.[Account], T2.[AcctName], T2.Levels

)

Select * from BSQuery

Union All

Select 3, 'NA', 'Profit/Loss for Period','', Sum(CurrentPeriod) from BSQuery

Order By 1,2
"




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
 foreach ($Row in 0..$RowCounts) 
{ 
   $powershell=[powershell]::Create()
   $powershell.RunspacePool=$Runspacepool

    $paramlist = @{
      fileRow=$ComPList
      count=$Row
      SqlStr=$SQL
       }
    $myScript={
        param($fileRow,
              $count,
              $hash,
              $SqlStr)
        #配置信息
        
        
        $Database   =  ($fileRow[$count] -split '\t')[1]
        $Serv     = ($fileRow[$count] -split '\t')[0]
        #$userName="montova"
        #$password="montova"
        #创建连接对象
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection


        #以 windows 认证连接 MSSQL
        $SqlConn.ConnectionString = "Data Source=$Serv;Initial Catalog=$Database;Integrated Security=SSPI;"

        
        #打开数据库连接

        $SqlConn.open()

        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlStr,$SqlConn)
         $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $SqlCmd
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataSet) |out-null  #将sql query的执行结果保存到adapter里面
   
            $hash[$count]=$dataset.Tables[0]          #[0] this is importandt key ,otherwise ,it will export-csv with system.dataset...system information not correct the value

           #$dataset.Tables >>'c:\temp\oufill.log'
            
            #关闭数据库连接
        $SqlConn.close()
       }
    
    [void]$powershell.AddScript($myScript).addargument($hash)
    [void]$powershell.AddParameters($paramlist)
    $handle=$powershell.BeginInvoke()
    $temp=''|select powershell,handle
    $temp.powershell=$powershell
    $temp.handle=$handle
    $jobs.Add($temp) | Out-Null
       
  }

  $return=$jobs|ForEach{ $_.powershell.endinvoke($_.handle) 
                     $_.powershell.dispose() }
  
  $jobs.Clear()
   $return     #it's important to get the result 
  $EndTime=Get-Date
 $hash.Keys |Sort-Object | ForEach-Object{  $hash[$_]  }  | export-csv C:\TEMP\invoic2incomingpayment4.csv -NoTypeInformation
  Write-Host -ForegroundColor Red ('running time is :' + ($EndTime-$starttime).totalseconds)

