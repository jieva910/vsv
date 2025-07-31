
taskkill /f /t /im   'SAP Business One.exe'


$UserName = 'jieva'
$Password = '0Ad+min001100002'

Function Test-ADAuthentication {
    param(
        $username,
        $password)
    
    (New-Object DirectoryServices.DirectoryEntry "",$username,$password).psbase.name -ne $null
}

$testAD = Test-ADAuthentication -username $UserName -password $password 
IF(!$testAD)
{Send-MailMessage -Body 'AD account password need to be changed' -From AUDIT@SAPB1APP.COM  -Subject 'AD password need to be changed' -To 'evan.ji@vesuvius.com'  -SmtpServer 'APMailrelay.vesuvius.com' -port 25
exit }


$adminUsername = 'CORP\jieva'                                      # 使用 pscredential 
$adminPassword = ConvertTo-SecureString $Password -AsPlainText -Force
$adminCreds = New-Object PSCredential $adminUsername, $adminPassword

Start-process 'C:\Program Files (x86)\SAP\SAP Business One\SAP Business One.exe' -credential $adminCreds -NoNewWindow 

Start-Sleep 10


$fullArgs= "powershell.exe -windowstyle hidden -ExecutionPolicy RemoteSigned -file C:\users\jievadm\Documents\ps\UI_EXPORT_GTS.ps1"

Start-process C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -credential $adminCreds -NoNewWindow -ArgumentList '-executionpolicy bypass', '-command',$fullArgs -WorkingDirectory C:\windows\SysWOW64

