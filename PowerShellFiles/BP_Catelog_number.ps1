
# BP catelog number for SZ
# 2023.04.07


$CMP = New-Object  -ComObject "SAPBOBSCOM.COMPANY"

. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

FN_CONNECTSAPB1 $CMP SZ

$oBPCatlog = $CMP.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oAlternateCatNum)


# $csv = Import-Csv 'C:\temp\JYXC BP Catelog.csv'  -Encoding UTF8

$csv = Import-Csv C:\temp\批量开票基础逻辑匹配表2023.csv -Encoding UTF8

foreach($r in $csv)
{

     IF ($oBPCatlog.GetByKey($r.ItemCode,$r.Cardcode,$r.ItemCode))
     {

      $oBPCatlog.UserFields.Fields.Item("U_VES_BPGTSItemName").value = $r.U_VES_BPGTSItemName

      $oBPCatlog.UserFields.Fields.Item("U_VES_BPGTSItemspec").value = $r.U_VES_BPGTSItemspec
       $oBPCatlog.UserFields.Fields.Item("U_VES_GTSUoM").value = $r.U_VES_GTSUoM
       $oBPCatlog.UserFields.Fields.Item("U_GTSComments").value = $r.U_GTSComments

      Write-Host $r.ItemCode $r.Cardcode "updated " $oBPCatlog.Update()        $cmp.GetLastErrorDescription()
     }
    <# ELSE 
     {
       $oBPCatlog.ItemCode =$r.ItemCode
       $oBPCatlog.CardCode = $r.Cardcode
       $oBPCatlog.Substitute = $r.ItemCode
        $oBPCatlog.UserFields.Fields.Item("U_VES_BPGTSItemName").value = $r.U_VES_BPGTSItemName
         $oBPCatlog.UserFields.Fields.Item("U_VES_BPGTSItemspec").value = $r.U_VES_BPGTSItemspec
               $oBPCatlog.UserFields.Fields.Item("U_VES_GTSUoM").value = $r.U_VES_GTSUoM
       $oBPCatlog.UserFields.Fields.Item("U_GTSComments").value = $r.U_GTSComments
      Write-Host $r.ItemCode $r.Cardcode "added "  $oBPCatlog.Add() $CMP.GetLastErrorDescription()

     } #>
}


#   SELECT  t.ItemCode,t.CardCode,t.Substitute,t.U_VES_BPGTSItemName,t.U_VES_BPGTSItemspec FROM oscn t

