
#using Runspacepool to update 340 PO's shipdate in line 
#2020.07.06

#-----------connect to SAPB1-------
	 $cmp = New-Object -COMObject 'SAPbobsCOM.Company'
	$cmp.Server ="SZ-SAPTST82" 
	$cmp.CompanyDB = "SAPB1_sz_TST"
	$cmp.DbServerType = 8
	$cmp.UserName = "jieva"
	$cmp.Password ="Ves-1234"
	$cmp.DbUserName="Butterfly"
    $cmp.DbPassword="buTterF1y"
	$cmp.LicenseServer = "sz-TSTSAPLIC92"
	[VOID]$cmp.Connect()
	if(-not $cmp.Connected) {Write-Host $cmp.GetLastErrorDescription() 
                              exit }
   
    # $cmp.XmlExportType = 3 
    $cmp.XMLAsString = 1
$xml =@'
<BOM>	<BO>		<AdmInfo>			<Object>22</Object>			<Version>2</Version>		</AdmInfo>		<Documents>			<row>				<DocEntry>$($docentry)</DocEntry>			</row>		</Documents>		<Document_Lines>			$($lines -join "`n")		</Document_Lines>	</BO></BOM>
'@

$docline=@'
<row> <LineNum>$($line.linenum)</LineNum> <ShipDate>$($line.shipdate.Replace( ".", ""))</ShipDate></row>
'@
# Create runspace session state
$InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

# Import all session functions into the runspace session state from the current one
Get-ChildItem Function:\ | Where-Object {$_.name -notlike "*:*"} |  select name -ExpandProperty name |
ForEach-Object {       
    # Get the function code
    $Definition = Get-Content "function:\$_" -ErrorAction Stop
    # Create a sessionstate function with the same name and code
    $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "$_", $Definition
    # Add the function to the session state
    $InitialSessionState.Commands.Add($SessionStateFunction)
}
#endregion

# --------------------------------------------------
#region - Runspace Pool Setup
# --------------------------------------------------
# Max runspace threads
$Limit = 10

# Create a Runspace Pool that uses the defined session state
$RunspacePool = [RunspaceFactory]::CreateRunspacePool($InitialSessionState)
[void]$RunspacePool.SetMinRunspaces(1) 
[void]$RunspacePool.SetMaxRunspaces($Limit) 
$RunspacePool.ApartmentState="MTA"
$PowerShell = [powershell]::Create() 
$PowerShell.RunspacePool = $RunspacePool   
$RunspacePool.Open()


$Hash =[hashtable]::Synchronized(@{}) 
$jobs = New-Object System.Collections.ArrayList

$starttime=get-date
function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}


$GroupsData=Import-Csv  -LiteralPath "\\sz-sz-sapb1app\SZ_PO_ShipDate\SZ_PO_NEW_SHIPDATE3.csv" |Group-Object DocEntry

foreach($grp in $GroupsData){
    $PowerShell = [powershell]::Create() 
    $PowerShell.RunspacePool = $RunspacePool   
    # Set parameters and values to import into the runspace that will be used in the script
    $ParamList = @{
        Param2 = $grp  
        SapCom = $cmp
        docline=$docline
        xml=$xml
          }        
    # Set the script to be run
    $MyScript = {
                  Param ( $Param2, 
                          $SapCom,$docline,$xml,
                          $hash)
                    $po=$SapCom.GetBusinessObject(22)
                     
                   
                       if ($po.getbykey($Param2.Name))
                        { $docentry=$Param2.Name
                          $lines=foreach($line in $Param2.Group){ $ExecutionContext.InvokeCommand.ExpandString($docline)   }
                          $fxml=$ExecutionContext.InvokeCommand.ExpandString($xml) 
                      
                          $po.Browser.ReadXml($fxml,0)
                          $err=$po.Update()
                       # $hash[$Param2.Name]  =$po.DocNum.ToString() +","+ $docentry.tostring() + " has been updated with Error code :"+$err.tostring() +" and error description :" +$SapCom.GetLastErrorDescription()	
                      $OUTPUT =$po.DocNum.ToString() +","+ $docentry.tostring() + " has been updated with Error code :"+$err.tostring() +" and error description :" +$SapCom.GetLastErrorDescription()	
                   
                        }
                   
                    
                  }	
                  #else{ $hash[$Param2.Name] =$Param2.Name +" PO entry not exist"}
                  #[pscustomobject]@{
                  #DocEntry = $err
                  #}
                       

    [void]$PowerShell.AddScript($MyScript).addargument($hash)
    [void]$PowerShell.AddParameters($ParamList)
    $Handle = $PowerShell.BeginInvoke()
    $temp = '' | Select PowerShell,Handle
    $temp.PowerShell = $PowerShell
    $temp.handle = $Handle        
    [void]$jobs.Add($Temp)        
 }

 $return = $jobs | ForEach {
    $_.powershell.EndInvoke($_.handle)
    $_.PowerShell.Dispose()
}
$jobs.clear()
#endregion
# --------------------------------------------------
#region - Return data
# --------------------------------------------------
$return 
#endregion
$endtime=get-date


<#save hashtable to file
 $outTime=Get-Date -format 'yyyyMMddHms'
 $outpath = "c:\temp\$outTime.log"
 $log= $hash.Keys | Sort-Object | ForEach-Object{ $hash[$_] }   
  set-Content $outpath -Value $log -Encoding Unicode
 #> 
Write-Host -ForegroundColor Red ('Total Runtime: ' + ($endtime - $starttime).TotalSeconds)

[void]$cmp.Disconnect()
Release-Ref($cmp)
