<#
  purpose: copy local csv file to remote server 
           call SAPB1 DI API version
           update data on the remote server
           sendmail to IT (optional)
#>

$remoteSvr='wg-wg-sapb1app'
$csvfilePath="C:\Temp\BPCODE.csv"
$Destionation_path = "C:\temp"
#$username = "$server\Administrator"
#$password =  Get-Content C:\mypassword.txt | ConvertTo-SecureString -AsPlainText -Force
#$creds = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password

$s = New-PSSession -ComputerName $remoteSvr -Credential corp\jievadm -ConfigurationName microsoft.powershell32 
Copy-Item -Path C:\Temp\BPCODE.csv -Destination $Destionation_path  -ToSession $s 

Invoke-Command -Session $s -FilePath C:\shared\PShell\psSAPB1\Update_BP_inactive.ps1