<#
   FUNCTION:  Call SAPBobs.com with Multiple Processs
   Date    :  2021.01
#>

[int]$Proc_Counter = 3                             # powershell Process number
$myScripts = @()                                    # Array of scriptblock
$csv='C:\Temp\tst_csitem.csv'                             # Csv file

$csvfileName = [System.IO.path]::GetFileNameWithoutExtension($csv)

# Define operation in SAPB1
$strUpdateItem = @‘ 
     $oITM=$cmp.getbusinessobject(4)
     write-host -ForegroundColor green Reading file $csvfile
    $CSV = Import-Csv $csvfile
    foreach ($x in $csv){      
          $cad=$x.ItemCode
         $t1=$x.U_Ves_PRP1
         $t2= $x.U_Ves_PRP2
        if ( $oITM.GetByKey($cad) ) {
            $oITM.UserFields.Fields.Item('U_Ves_PRP1').value = $t1
            $oITM.UserFields.Fields.Item('U_Ves_PRP2').value =$t2
            $outputcode = $oITM.Update()
            write-host  $x.ItemCode updated with error code $outputcode
            }
          }

’@

# Process 1 connection
$strConnCom1 = @‘
  param($csvfile)
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
$cmp.Server = 'SZ-SAPtst82'
$cmp.CompanyDB ='SAPB1_cs_tst'
$cmp.DbServerType = 8
$cmp.UserName = 'jieva'
$cmp.Password ='vesint99'
# $cmp.DbUserName=$cmpdbuser
# $cmp.DbPassword=$cmpdbpwd
$cmp.UseTrusted=$true
$cmp.LicenseServer = 'SZ-TSTSAPLIC92'
[void]$cmp.Connect()      
if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;break             } 
else { Write-Host -ForegroundColor Cyan $cmp.UserName connected successfully}
’@


# Process 2 connection
$strConnCom2 = @‘
  param($csvfile)
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
$cmp.Server = 'SZ-SAPtst82'
$cmp.CompanyDB ='SAPB1_cs_tst'
$cmp.DbServerType = 8
$cmp.UserName = 'wangdol'
$cmp.Password ='Ves-123456'
# $cmp.DbUserName=$cmp2dbuser
# $cmp.DbPassword=$cmp2dbpwd
$cmp.UseTrusted=$true
$cmp.LicenseServer = 'SZ-TSTSAPLIC92'
[void]$cmp.Connect()  
if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;break            } 
else { Write-Host -ForegroundColor Cyan $cmp.username connected successfully}
’@


# Process 3 connection
$strConnCom3 = @‘
  param($csvfile)
$cmp = New-Object -COMObject 'SAPbobsCOM.Company'
$cmp.Server = 'SZ-SAPtst82'
$cmp.CompanyDB ='SAPB1_cs_tst'
$cmp.DbServerType = 8
$cmp.UserName = 'jinman'
$cmp.Password ='Ves-123456'
# $cmp.DbUserName=$cmp2dbuser
# $cmp.DbPassword=$cmp2dbpwd
$cmp.UseTrusted=$true
$cmp.LicenseServer = 'SZ-TSTSAPLIC92'
[void]$cmp.Connect()  
if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() ;break            } 
else { Write-Host -ForegroundColor Cyan $cmp.username connected successfully}
’@

$strScriptblock1 = $strConnCom1,$strUpdateItem 
$strScriptblock1 -join "\n"

$strScriptblock2 = $strConnCom2,$strUpdateItem 
$strScriptblock2 -join "\n"

$strScriptblock3 = $strConnCom3,$strUpdateItem 
$strScriptblock3 -join "\n"

# dynamicly create scriptblock with current string variables
Get-Variable -Name 'strScriptblock*' | % {  $myScripts += [scriptblock]::Create($_.Value) }

if ($myScripts.Length -lt $Proc_Counter) {write-host -ForegroundColor red  ("process scriptblocks: {0}`r`nprocess counter: {1}`r`nprocess scriptblocks must >= process counter !" -f $($myScripts.Length),$($Proc_Counter)) ;exit}

# Function -Split File 使用regex 搜索，但是 最后一行有问题,还没找到运行
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
             if ($count )
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


function fn_SplitCSV
{
  Param($sourceCSV,$num,$csvfileName)
  
    $TotalRows = (Import-CSV $sourceCSV ).count

    [INT]$rows_perfile = $TotalRows / $num    # 定义生成的文件个数

    # variable used to advance the number of the row from which the export starts
    $startrow = 0 ;

    # counter used in names of resulting CSV files
    $counter = 1 ;

    # setting the while loop to continue as long as the value of the $startrow variable is smaller than the number of rows in your source CSV file
    while ($startrow -lt $TotalRows)
    {
        # import of however many rows you want the resulting CSV to contain starting from the $startrow position and export of the imported content to a new file
        Import-CSV $sourceCSV | select-object -skip $startrow -first $rows_perfile | Export-CSV "c:\temp\$($csvfileName)_$($counter).csv"  -NoTypeInformation ;

        # advancing the number of the row from which the export starts
        $startrow += $rows_perfile 

        # incrementing the $counter variable
        $counter++ ;
     }

}
fn_SplitCSV $csv $Proc_Counter $csvfileName

$starttime = Get-Date

$loop = 1 
foreach($mys in $myScripts)
{
  $newCSVfile = "c:\temp\$($csvfileName)_$loop.csv"
  Start-Process powershell "&{$mys} $newCSVfile"              # notice this sign & scriptblock + argument
  if ($loop -eq $Proc_Counter) {break}                        # if reach process counter then jump to loop
  $loop++
}

$endtime =Get-Date
Write-output ('Running Time is :' + ($endtime-$starttime).TotalSeconds)>>"c:\temp\tst_csitem.log"

