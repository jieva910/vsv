# Purpose : export licensed user in site list
# Date    : 2020.11

cls

$licenseServer =  "sz-saplic92"
$Dest   = "C:\Temp\B1LicXML\" 

del "$Dest\*.xml"

$SourceXML88= "C:\Program Files (x86)\SAP\SAP Business One ServerTools\License\B1Upf.xml"
$SourceXML92 = "C:\Program Files (x86)\SAP\SAP Business One ServerTools\LicenseHTTPS\webapps\lib\B1Upf.xml"


$Username = "corp\sapdtwserv"
$Password = ConvertTo-SecureString "v3sSL2.b1" -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential($Username, $Password)


Function Test-ADAuthentication {
    param(
        $username,
        $password)
    
    (New-Object DirectoryServices.DirectoryEntry "",$username,$password).psbase.name -ne $null
}

$testAD = Test-ADAuthentication -username $UserName -password "v3sSL2.b1"
IF(!$testAD)
{ write-host 'user password not correct' ;break }

Write-Host -ForegroundColor Cyan " copy SAPB1 license xml file from server to local ...."
foreach ( $svr in $licenseServer )
{
    $Session = New-PSSession -ComputerName $svr -Credential $mycreds
   
    if ( $svr -in    "sz-saplic92","dg-saplic01" ,"dg-saplic02" )
    {  
        
     Copy-Item $SourceXML92 -Destination "$Dest$svr-B1Upf.xml" -FromSession $Session
      
    }
    else { Copy-Item $SourceXML88 -Destination "$Dest$svr-B1Upf.xml" -FromSession $Session }
   
}

 
 Write-Host 
 Write-Host -ForegroundColor Cyan " Read all SAPb1 xml in folder "


$XMLfiles = Get-ChildItem -Path "C:\Temp\B1LicXML\*.xml" 

 Write-Host 
 Write-Host -ForegroundColor Cyan " Retrieve Users from all XML file "

$User1ist = foreach ($xmlfile in $XMLfiles)
{
    $xmlPath = $xmlfile.FullName 
    if (Test-Path -Path $xmlPath)
    { 
      [XML]$xmlDoc = Get-Content $xmlPath

      # Using XML XPath to retrieve user 
      # Select-Xml '//User[contains(Modules/Module/KeyType, "PROFESSIONAL")]' $xmlDoc |%{$_.Node.UserName}

    # get user who assigned SAPB1 Professional license
      $xmlDoc.SelectNodes("/Users/User[Modules/Module/KeyType = 'PROFESSIONAL']")|%{$_.UserName}

    }
}

# LDAP Query user AD
Function Fn_Return_AD_Usr_infor ($sUsrName)
{
   $Searcher = New-Object DirectoryServices.DirectorySearcher
        $Searcher.SearchRoot = 'LDAP://DC=corp,DC=vesuvius,DC=com'
        $Searcher.Filter = '(&(objectCategory=user)(cn='+$sUsrName+'))'
        $res = $Searcher.FindAll()  | Sort-Object path
   $Value = "" | Select-Object -Property User_code,first_name ,last_name,email,displayName,phonenum,empid
        
        foreach ($usrTmp in $res)
        {  $Value.User_code = $usrTmp.Properties["name"]
           $Value.first_name = $usrTmp.Properties["sn"] 
            $Value.last_name  = $usrTmp.Properties["givenname"]
            $Value.email = $usrTmp.Properties["mail"] 
            $Value.displayName= $usrTmp.Properties["displayName"]
            $Value.phonenum=$usrTmp.Properties["telephoneNumber"]
            $Value.empid=$usrTmp.Properties["employeeid"]

          #$usrtmp.Properties
        }

  Return $Value
}


 Write-Host 
 Write-Host -ForegroundColor Cyan " Query users in Active Directory and get the user who has no AD account ... "

 $users = foreach ( $usr in $User1ist ) 
{ 
   $v = Fn_Return_AD_Usr_infor $usr
  if ($v.User_code -eq $null)             # Retrieve user who has no AD information
   { $usr }
 
 }

 Write-Host 
 Write-Host -ForegroundColor Cyan " Convert User list to string and used in SQL query... "

[array]$UserArray =   $User1ist | foreach { "'" + $_ + "'" }

[string]$UserString = $null

 $UserString = $UserArray -join ","     # this varible will be used in "C:\shared\PShell\pshell_sql\query_sapb1LicensedUser.ps1"


  Write-Host 
 Write-Host -ForegroundColor Cyan " start another PS file to excute sql query in all sapb1 company dbs... "

$scriptPath = "C:\shared\PShell\pshell_sql\query_sapb1LicensedUser.ps1"  # 原版 for SAPB1 8.8


  $SLDServs= "sz-saplic92"             # varible will be reference in below ps1 file ,,,,... dg-sapsld10
 
 $scriptPath2 = "C:\shared\PShell\pshell_sql\Get_dbs_from_SLDSVR.ps1"    # 改进版本


 Invoke-Expression "& $scriptPath2"

