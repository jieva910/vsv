<#
   FUNCTION:  Call SAPBobs.com with Multiple Processs
   Date    :  2021.01
#>

[int]$Proc_Counter = 3                             # powershell Process number
$myScripts = @()                                    # Array of scriptblock
$csv='C:\Temp\SZBP.csv'                             # Csv file

$csvfileName = [System.IO.path]::GetFileNameWithoutExtension($csv)

# Define operation in SAPB1
$strUpdateBP = @‘
    $starttime = Get-Date
     $oBP=$cmp.getbusinessobject(2)
    $CSV = Import-Csv $csvfile
    $Results = foreach ($x in $csv){      
          $cad=$x.CardCode
         $t1=$x.CreditLine
         $t2= $x.DebtLine
        if ( $oBP.GetByKey($cad) ) {
            $oBP.CreditLimit = $t1
            $oBP.MaxCommitment =$t2
            $outputcode = $oBP.Update()
           [PSCustomobject] @{ 
                Success = $outputcode       # 假设传入参数5时失败，其余成功
                cardcode = $cad
                ThreadId = [AppDomain]::GetCurrentThreadId()   # 当前线程ID  
            }
          }
    }
 Tee-Object  -InputObject $Results   "c:\temp\BPcreditlimite.log" -Append
$endtime =Get-Date
Write-output ('Running Time is :' + ($endtime-$starttime).TotalSeconds)>>"c:\temp\BPcreditlimite.log"
’@

# Process 1 connection
$strConnCom1 = @‘
  param($csvfile)
$starttime = Get-Date
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
$cmp.Server = 'SZ-SAPtst82'
$cmp.CompanyDB ='SAPB1_SZ_tst'
$cmp.DbServerType = 8
$cmp.UserName = 'jieva'
$cmp.Password ='Ves-1234'
# $cmp.DbUserName=$cmpdbuser
# $cmp.DbPassword=$cmpdbpwd
$cmp.UseTrusted=$true
$cmp.LicenseServer = 'SZ-TSTSAPLIC92'
[void]$cmp.Connect()      
if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;break             } 
else { Write-Host -ForegroundColor Cyan $cmp.CompanyDB connected successfully}
’@


# Process 2 connection
$strConnCom2 = @‘
  param($csvfile)
$starttime = Get-Date
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
$cmp.Server = 'SZ-SAPtst82'
$cmp.CompanyDB ='SAPB1_SZ_tst'
$cmp.DbServerType = 8
$cmp.UserName = 'shenste'
$cmp.Password ='Ves-123456'
# $cmp.DbUserName=$cmp2dbuser
# $cmp.DbPassword=$cmp2dbpwd
$cmp.UseTrusted=$true
$cmp.LicenseServer = 'SZ-TSTSAPLIC92'
[void]$cmp.Connect()  
if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;break            } 
else { Write-Host -ForegroundColor Cyan $cmp.username connected successfully}
’@

$strScriptblock1 = $strConnCom1,$strUpdateBP 
$strScriptblock1 -join "\n"

$strScriptblock2 = $strConnCom2,$strUpdateBP 
$strScriptblock2 -join "\n"


# dynamicly create scriptblock with current string variables
Get-Variable -Name 'strScriptblock*' | % {  $myScripts += [scriptblock]::Create($_.Value) }

if ($myScripts.Length -lt $Proc_Counter) {write-host -ForegroundColor red  ("process scriptblocks: {0}`r`nprocess counter: {1}`r`nprocess scriptblocks must >= process counter !" -f $($myScripts.Length),$($Proc_Counter)) ;exit}

# Function -Split File
Function Split-CSV ($csv,$Proc_Counter,$csvfileName) {
    $count = 1
    $fileRows = (Get-Content  $csv -ReadCount 0 ).count 
    $header = get-content $csv -TotalCount 1
    [int]$readcount = $fileRows/$Proc_Counter
    get-content $csv -ReadCount $readcount |               # define new CSV rows
      foreach {
           #add tail entries from last batch to beginning of this batch
           $newbatch = $tail + $_ 

           #create regex to match last entry in this batch
           $regex = '^' + [regex]::Escape(($newbatch[-1].split(',')[0])) 

           #Extract everything that doesn't match the last entry to new file

             #Add header if this is not the first file
             if ($count)
               {
                 $header |
                   set-content "c:\temp\$($csvfileName)_$count.csv"
                }
             $newbatch -notmatch $regex | 
              add-content "c:\temp\$($csvfileName)_$count.csv"  

           #Extact tail entries to add to next batch
           $tail = @($newbatch -match $regex)

           #Increment file counter
           $count++ 
      }
}

Split-CSV $csv $Proc_Counter $csvfileName


$loop = 1 
foreach($mys in $myScripts)
{
  $newCSVfile = "c:\temp\$($csvfileName)_$loop.csv"
  Start-Process powershell "&{$mys} $newCSVfile"              # notice this sign & scriptblock + argument
  if ($loop -eq $Proc_Counter) {break}                        # if reach process counter then jump to loop
  $loop++
}