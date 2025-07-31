

$cmp  = new-object -ComObject "sapbobscom.company"

$cmp.Server = "SZ-SAP01"
$cmp.SLDServer ="SZ-SAPLIC92:40000"
$cmp.CompanyDB = "SAPB1-SZ"
$cmp.DbServerType = 8
$cmp.UseTrusted = 1 
$cmp.UserName="\"
$cmp.Password=""

$cmp.Connect()
$cmp.GetLastErrorDescription()



$targetTime = Get-Date "12:10 PM"

# Loop until the current time reaches or passes the target time
while ((Get-Date) -lt $targetTime) {
    # Wait for 1 second before checking again (to avoid overloading the CPU)
    Start-Sleep -Seconds 1
}


 $oUF2 = $cmp.GetBusinessObjectFromXML("C:\Dell\33.XML",0) 
  $oUF2.add()   
# When the target time is reached, run your command
"Add 33 UDFs $($cmp.GetLastErrorDescription())" | Out-File -FilePath "C:\temp\TimeLog.txt" -Append

 $oUF3 = $cmp.GetBusinessObjectFromXML("C:\Dell\34.XML",0) 
  $oUF3.add()  

  "Add 34 UDFs $($cmp.GetLastErrorDescription())" | Out-File -FilePath "C:\temp\TimeLog.txt" -Append

   $oUF4 = $cmp.GetBusinessObjectFromXML("C:\Dell\35.XML",0) 
  $oUF4.add()  

  "Add 35 UDFs $($cmp.GetLastErrorDescription())" | Out-File -FilePath "C:\temp\TimeLog.txt" -Append

  
   $oUF5 = $cmp.GetBusinessObjectFromXML("C:\Dell\36.XML",0) 
  $oUF5.add()  

  "Add 36 UDFs $($cmp.GetLastErrorDescription())" | Out-File -FilePath "C:\temp\TimeLog.txt" -Append

