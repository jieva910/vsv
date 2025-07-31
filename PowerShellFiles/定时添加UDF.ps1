


# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

$site = "cstst"

  Fn_ConnectSAPB1 $site

# export UDF of RDR1 table to xml
 $UDF = $CMP.GetBusinessObject(152)    ##152 UserFieldsMD object 
 $CMP.XmlExportType = 3 
# TableID	FieldID	 AliasID
# RDR1	     57	     Ves_RowComment
# query table CUFD
 $UDF.GetByKey('RDR1',57)
 $UDF.SaveXML('C:\TEMP\VES_ROWCOMMENT.XML')


 function add_UDF 
{
 . C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1
  $site = "cs"
  Fn_ConnectSAPB1 $site

  $udf_new = $cmp.GetBusinessObjectFromXML('C:\TEMP\VES_ROWCOMMENT.XML',0)

  Write-Host 'UDF has been added with errcode:'   $udf_new.Add() $cmp.GetLastErrorDescription()
 

}

$TriggerTimes = @(
    '05:40:59pm'
  #  '10:16:59am'
   )

# Sort in chronologic order
#  assuming the times format are the same
$TriggerTimes = $TriggerTimes | Sort-Object


foreach ($t in $TriggerTimes)
{
    # Past time ?
    if((Get-Date) -lt (Get-Date -Date $t))
    {
        # Sleeping
        while ((Get-Date -Date $t) -gt (Get-Date))
        {
            # Sleep for the remaining time
            (Get-Date -Date $t) - (Get-Date) | Start-Sleep
        }

        # Trigger event
        #  insert your code here
        "# TriggerTime: '$t' - Executing my code here!"

       add_UDF

    }else{"Belong to the past: '$t'"}
}

