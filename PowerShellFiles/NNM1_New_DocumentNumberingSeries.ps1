
<#
  Purpose : Add New Document Numbering Series for New Posting Periods
  Date    : 2020.12
#>


cls

$SourceSite    = "CSTST"
$cmp = New-Object -ComObject "sapbobscom.company"
# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp $SourceSite


# sql query existing document numbering series
# save recordset as csv
$oRs = $cmp.GetBusinessObject(300)

$sql ="select ObjectCode, replace(SeriesName,RIGHT(SeriesName,2),right(CAST(YEAR(GETDATE())+1 as NVARCHAR(4)),2))
,InitialNum 
,LastNum 
,GroupCode,DocSubType
,isnull(Remark,'') Remark
,year(getdate())+1 Indicator from nnm1
WHERE Indicator = CAST(year(getdate()) AS NVARCHAR(4))
"

$oRs.DoQuery($sql)  

$hash = @{}

$RecordData = while ($oRs.EOF -ne $True)
{
    foreach ($field in $ors.Fields)
    {
       $hash.$($field.name) = $field.value
    
    }

    #$hash
    #New-Object PSObject -property $hash

    [PSCustomObject]$hash
    $oRs.MoveNext()
}

$RecordData |Export-Csv -LiteralPath C:\Temp\DocNumberingSeries.csv -NoTypeInformation


# Read new document numbering series in CSV file and Add them 

$oCmpSrv = $cmp.GetCompanyService()
$oSeriesService = $oCmpSrv.GetBusinessService(35)  #   ServiceTypes.SeriesService
$oSeries = $oSeriesService.GetDataInterface(0)     #   ssdiSeries

$csvDocNumbering = Import-Csv -LiteralPath C:\Temp\DocNumberingSeries.csv
foreach($row in $csvDocNumbering)
{
     $oSeries.Document     = $row.ObjectCode
    # series name
    $oSeries.Name          = $row.SeriesName
    # the first number
    $oSeries.InitialNumber = $row.InitialNum
    # last number
    $oSeries.LastNumber    = $row.LastNum
    # the group code
    $oSeries.GroupCode     = $row.GroupCode
    # DocumentSubType
     $oSeries.DocumentSubType = $row.DocSubType
     # remark
     $oSeries.Remarks         = $row.Remark
    # the period indicator
    $oSeries.PeriodIndicator  = $row.Indicator
    
    #before adding the series to the document ,check that the lastNumber property
    #of the last series has a value(if not the add function will fail)
    
    Write-Host $oSeries.Name   $oSeriesService.AddSeries($oSeries) $cmp.GetLastErrorDescription()

}

