

$cmp = New-Object -ComObject "sapbobscom.company"

. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

Fn_ConnectSAPB1 $cmp “SZTST"


$oJe = $cmp.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oJournalEntries)
$oRs = $cmp.GetBusinessObject(300)


$sql_JE = "  SELECT NUMBER , replace(replace(CAST(U_Ves_JE_Attach AS NVARCHAR(1000)),'\\SZ-NAS02\SAPAttachments','\\SZ-FS02\SAPAttachments_Hist'),'\\SZ-NAS01\SAPAttachments','\\SZ-FS02\SAPAttachments_Hist') 'newpath' FROM OJDT t WHERE t.ObjType = 30 
  AND  t.U_Ves_JE_Attach IS NOT NULL AND t.U_Ves_JE_Attach LIKE '%nas%'"

  $oRs.DoQuery($sql_JE)

  if (!$oRs.EoF) {
  
  
   [xml]$xml = $ors.GetAsXML()
   $Nodes = $xml.SelectNodes("//row")

   foreach($n in $nodes)
    {
      if ($oJe.GetByKey($n.Number))
      { $oJe.UserFields.Fields.Item("U_Ves_JE_Attach").value = $n.newpath.tostring()
        
        write-host $n.Number  $oJe.Update() $cmp.GetLastErrorDescription()
      }
    }

    }






$opo  = $cmp.GetBusinessObject(22)

$sql_PO = " select  t.docentry,replace(replace(CAST(U_Ves_linkAttach AS NVARCHAR(1000)),'\\SZ-NAS02\SAPAttachments','\\SZ-FS02\SAPAttachments_Hist'),'\\SZ-NAS01\SAPAttachments','\\SZ-FS02\SAPAttachments_Hist') FROM opor
 t WHERE  t.U_Ves_linkAttach IS NOT NULL AND t.U_Ves_linkAttach LIKE '%nas%'"


$opo.GetByKey(13)

$opo.UserFields.Fields.Item("U_Ves_LinkAttach").value = "\\SZ-FS02\SAPAttachments_Hist\20160705105924565_0001.pdf"
$opo.Update()
$cmp.GetLastErrorDescription()