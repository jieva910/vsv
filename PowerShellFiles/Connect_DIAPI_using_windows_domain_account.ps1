$cmp = New-Object -ComObject "SAPBOBSCom.company"

$cmp.SLDServer = 'https://sz-tstsaplic92:40000'
$cmp.Server = 'SZ-SAPTST82'
$cmp.CompanyDB = 'SAPB1_CS_TST'
$cmp.DbServerType = 8
$cmp.UseTrusted = $true
$cmp.UserName = '\'
 $cmp.Password =''
$cmp.Connect()
$cmp.GetLastErrorDescription()