# 2022.10.12
# update bank information of employee supplier 

$cmp = New-Object -ComObject 'sapbobscom.company'
$site = 'SZ'

. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

fn_connectsapb1 $cmp $site


$oSupplier = $cmp.GetBusinessObject(2)

$csv = Import-Csv C:\Temp\ASemployeesupplier.csv ";"

foreach($v in $csv)
{
if ($oSupplier.GetByKey($v.SAP))
{

  $oSupplier.BPBankAccounts.BankCode = 'CMB'
$oSupplier.BPBankAccounts.AccountNo =$v.bankno
  $oSupplier.BPBankAccounts.AccountName =$v.员工姓名
  $oSupplier.BPBankAccounts.Branch ='update'
    $oSupplier.BPBankAccounts.City = '鞍山'
     $oSupplier.BPBankAccounts.UserNo1 = $v.bank
}


$v ;$oSupplier.Update();$cmp.GetLastErrorDescription()
}


# IMPORT YK Supplier master data
$oSupplier = $cmp.GetBusinessObject(2)

$csvYKBP = Import-Csv C:\Temp\YKBP.csv

foreach($bp in $csvYKBP)
{
 $oSupplier.CardCode = $bp.bpcodee
 $oSupplier.CardName =$bp.业务伙伴名称
 $oSupplier.Currency ='RMB'
 $oSupplier.CardType = [SAPbobsCOM.BoCardTypes]::cSupplier
 $oSupplieR.Phone1 = $bp.'电话 1'
 $oSupplier.ShippingType = $bp.'Shipping Type'
$oSupplier.ContactEmployees.Name = $bp.ContactPerson
$oSupplier.Addresses.AddressType = [SAPbobsCOM.BoAddressType]::bo_BillTo
$oSupplier.Addresses.AddressName = $bp.AddressName
$oSupplier.Addresses.county = $bp.业务伙伴名称
$oSupplier.Addresses.Country = 'CN'
$oSupplier.Addresses.Street = $bp.'Street/PO Box'
$oSupplier.PeymentMethodCode = 'S-T/T'
$oSupplier.BPPaymentMethods.PaymentMethodCode = 'S-T/T'
$oSupplier.BPPaymentMethods.SetCurrentLine(0)

$oSupplier.PayTermsGrpCode = $bp.'Payment Terms'
$oSupplier.BankCountry = 'CN'
$oSupplier.BPBankAccounts.BankCode =$bp.银行代码
$oSupplier.BPBankAccounts.AccountName = $bp.'Bank Account Name'
$oSupplier.BPBankAccounts.AccountNo = $bp.'Bank Account'
$oSupplier.BPBankAccounts.Branch = $bp.'Bank Branch'
$oSupplier.BPBankAccounts.UserNo1=$bp.'用户编号 1'
$oSupplier.BPBankAccounts.City = $bp.城市
$oSupplier.BPBankAccounts.State = 11
$oSupplier.VatLiable = 1  
$oSupplier.VatGroup = $bp.税收组
$oSupplier.Properties(1) = 1 
$oSupplier.GroupCode = 103

Write-Output  $bp.bpcodee $oSupplier.ADD() $cmp.GetLastErrorDescription()
}


# deactivate employee supplier master data

$oEmployeeSupplierCode = $cmp.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oBusinessPartners)
$file= Import-Csv C:\dell\EMP_SUPPLIER.csv

foreach($n in $file)
{
  if($oEmployeeSupplierCode.GetByKey($n.SUPPLIERCODE))
  {   
   $oEmployeeSupplierCode.Frozen = 1
   $oEmployeeSupplierCode.Valid = 0 
   $oEmployeeSupplierCode.FrozenRemarks = "INC0394778"
    Write-host $n.SUPPLIERCODE $oEmployeeSupplierCode.update() $cmp.GetLastErrorDescription()
  } 
}