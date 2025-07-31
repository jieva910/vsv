cls

$ticknum = '9999'

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$site = "cstst"

  Fn_ConnectSAPB1 $site

   # Disable SP control 
  fn_SAPB1_SP_control $ticknum 'N'  $site 


  $bp=$cmp.GetBusinessObject(2)


 # 1 以下可以在本机 32bit  DIAPI 运行，CS TST SAP DB,如果连接CK SAPB1 运行会报‘Cannot convert null to type "System.DateTime",
 # 检查后发现是由于ISE 或者 powershell.exe 版本不一致，本机ISE版本为10， DG-SAPSTG91上的ISE 版本为 6
  $csv = Import-Csv C:\Temp\CKBP.csv

  foreach($r in $csv){
  if ($bp.GetByKey($r.'BP Code')){  
      if(![string]::IsNullOrEmpty($r.'Active From')){$bp.ValidFrom=$null}
      if(![string]::IsNullOrEmpty($r.'Active To')){$bp.ValidTo=$null}
    Write-Host "BP:"$r.'BP Code' " updated with err code:" $bp.Update() $cmp.GetLastErrorDescription()
    }
  }



  # 以下是验证 datetime 可以设置为 $null
  try { [datetime]$null; write-output 'worked!' } catch { write-output 'no worked!' }
try { [nullable[datetime]]$null; write-output 'worked!' } catch { write-output 'no worked!' }






   # 以下 使用 xml能够将 BP主数据上的 活跃日期设置为 空
  $bp.GetByKey('SO47482test9')
  $cmp.XmlExportType = 3 
  $cmp.XMLAsString = 1 
  [xml]$xmlitm = $bp.GetAsXML()
   $nodes = $xmlitm.SelectNodes('//BusinessPartners/row')

   $nodes | %{
     $node_validfrom = $_.SelectSingleNode('ValidFrom')
      $node_validto = $_.SelectSingleNode('ValidTo')
       if ($node_validfrom) {[void]$_.removechild($node_validfrom) ; $result = $true}
       if ($node_validto) {[void]$_.removechild($node_validto) ; $result = $true}
   if($result){
       $bp2 = $cmp.GetBusinessObjectFromXML($xmlitm.InnerXml,0)
       $bp2.Update()
       $cmp.GetLastErrorDescription()
       }
 
   }



   # Disable SP control 
   $ticknum=''
  fn_SAPB1_SP_control $ticknum 'Y'  $site 



  # FOR LIVE 

    
cls

$SourceSite    = "kt"

$ticktNum = 'INC0195792 '

# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$CMP = New-Object -ComObject "SAPBOBSCOM.COMPANY"
# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $CMP $SourceSite

fn_SAPB1_SP_control $ticktNum 'N' $SourceSite

$oBP = $cmp.GetBusinessObject(2)

$ors = $cmp.GetBusinessObject(300)
$sql = "SELECT CardCode,validfor,validfrom,validto FROM OCRD T0 WHERE validfor= 'y' and ( isnull(validfrom,'')<>'' or isnull(validto,'')<>'') and validto >'20220101' and (validfrom <>validto)"

$ors.DoQuery($sql)

if (!$ors.EoF)
{
 [xml]$xmls = $ors.GetAsXML()
 $nodes = $xmls.SelectNodes("//row")
 foreach($n in $nodes)
 {
  if($oBP.GetByKey($n.CardCode))
  {
    if(![string]::IsNullOrEmpty($oBP.ValidFrom)){$oBP.ValidFrom =$null}
    if(![string]::IsNullOrEmpty($oBP.ValidTo)){$obp.ValidTo = $null}
    Write-Host "$($n.CardCode) validfrom validto set null $($oBP.Update()) $($cmp.GetLastErrorDescription())"
  }
 }
}


