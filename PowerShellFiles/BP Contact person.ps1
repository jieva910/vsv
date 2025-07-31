$cmp = New-Object -COMObject 'SAPbobsCOM.Company'


. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

FN_CONNECTSAPB1 $cmp "CS"

$csv = Import-Csv C:\dell\cs.csv -Delimiter "," 

# add contact person 
foreach($r in $csv)
 {
    $obp = $cmp.GetBusinessObject(2)

    if($obp.GetByKey($r.BP))
    {  
       
    $i = $obp.ContactEmployees.Count
    $obp.ContactEmployees.Add()

 
    $obp.ContactEmployees.SetCurrentLine(0)
     if (!$obp.ContactEmployees.Name){  $obp.ContactEmployees.SetCurrentLine(0)    }
     else {$obp.ContactEmployees.SetCurrentLine($i)}
       $obp.ContactEmployees.Name = 'E_Invoice to Customer'
       $obp.ContactEmployees.E_Mail = $r.'E_Invoice to Customer'
     write-host "$($r.BP) has been added with err code: "  $obp.update() $cmp.GetLastErrorDescription()
       }
 }


 # update contact person name  
foreach($r in $csv)
 {
    $obp = $cmp.GetBusinessObject(2)

    if($obp.GetByKey($r.BP))
    {  
       
    $i = $obp.ContactEmployees.Count
   
    for($j = 0;$j -lt $i;$j++)
    {
    $obp.ContactEmployees.SetCurrentLine($j)
     if ($obp.ContactEmployees.Name -eq 'E_Invoice to Customer')
     {  
       $obp.ContactEmployees.E_Mail = $r.'E_Invoice to Customer'
     write-host "$($r.BP) has been updated  with err code: "  $obp.update() $cmp.GetLastErrorDescription()
     break   }
     }
     
       }
 }


 $csv = import-csv C:\Temp\csbp.csv

   $oBP=$cmp.getbusinessobject(2)

   foreach($c in $csv)
   {
    if ( $oBP.GetByKey($c.CardCode) ) {
        $obp.Properties(42) = $c.pro42 
       $obp.Properties(43) = $c.pro43 
        $obp.Properties(44) = $c.pro44 
        
        
        $outputcode = $oBP.Update()
        $outlog =$cmp.CompanyDB + ' '+ $c.CardCode + ' updated  with error code:' +   $outputcode + ' and error description is:' + $cmp.GetLastErrorDescription()
        $outlog
       }
   }
    