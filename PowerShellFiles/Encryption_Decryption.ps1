# encrypt user password 

$EncryptionKeyBytes = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($EncryptionKeyBytes)
$EncryptionKeyBytes | Out-File "c:\temp\encryption.key"

$EncryptionKeyData = Get-Content "c:\temp\encryption.key"

 $content = Read-Host -AsSecureString 
# $content = Get-Content C:\Temp\tst.ps1

$content | ConvertFrom-SecureString -Key $EncryptionKeyData | Out-File -FilePath "c:\temp\secret.encrypted"


$EncryptionKeyData = Get-Content "c:\temp\encryption.key"
$PasswordSecureString = Get-Content "c:\temp\secret.encrypted" | ConvertTo-SecureString -Key $EncryptionKeyData
$PlainTextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordSecureString))

 $PlainTextPassword

# encrypt ps file 2022.11

$Code = Get-Content "C:\temp\tst.txt" 
$CodeSecureString = ConvertTo-SecureString  $Code -AsPlainText -Force
$Encrypted = ConvertFrom-SecureString -SecureString $CodeSecureString
$Encrypted | Out-File -FilePath 'C:\temp\tst_encrpt.txt'


$Instructions = Get-Content C:\temp\tst_encrpt.txt
$Decrypt = $Instructions | ConvertTo-SecureString
$Code = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Decrypt))
Invoke-Expression $Code