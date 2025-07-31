 <# $CMP = New-Object  -ComObject "SAPBOBSCOM.COMPANY"

  . C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

    FN_CONNECTSAPB1 $CMP "sztst"

    $oEmp = $CMP.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oEmployeesInfo)

    $oTeam  = $CMP.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oTeams)

    $oemp.GetByKey(229)


   $oTeam.GetByKey(2)   # Teams ID of Employee master data
   for ( $i= 1 ;$i -lt $oteam.TeamMembers.Count;$i++){
     $oTeam.TeamMembers.SetCurrentLine($i)

    if ( $oTeam.TeamMembers.EmployeeID =144)
    { 
      $oTeam.TeamMembers.EmployeeID = 1
      $oTeam.TeamMembers.RoleInTeam = 0
      Write-Host $oTeam.TeamMembers.EmployeeID   $oTeam.UPDATE() $CMP.GetLastErrorDescription()
       break
    } 
   }


   # DI API 不支持删除EMPLOYEE'S TEAM ID 2023.04
   
   $oTeam.GetByKey(2) 
   $oteam.TeamMembers.Add()
   $oteam.TeamMembers.EmployeeID = 144
    
      Write-Host  $oTeam.UPDATE() $CMP.GetLastErrorDescription()
#>

     

     # get user list which locked then set employee master to inactive
function set_inacive_employee 
{
     $sql = "SELECT t1.empID FROM OUSR t INNER JOIN OHEM t1 ON t.userId = t1.userId WHERE t.Locked = 'Y' AND t.GROUPS <> 99"
     $oRs = $CMP.GetBusinessObject(300)
     $oEmp = $CMP.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oEmployeesInfo)
     $oRs.DoQuery($sql)
    $xmlfile = [xml]$ors.GetAsXML()
    $Nodes = $xmlfile.SelectNodes("//row")

    foreach($nd in $Nodes)
    {   
    
     if ($oemp.GetByKey($nd.empID))
     {   $oemp.Active = 0 
       Write-Host $oEmp.EmployeeID $oEmp.Update() $CMP.GetLastErrorDescription()
      
     }
    
    
    }
     
}

     $sites = 'AS','BY','CS','HG','KT','SQ','SZ','WE','WN','YK'
$ticktNum ="INC0344888"

  foreach($s in $sites)
  {
  
    
   $CMP = New-Object  -ComObject "SAPBOBSCOM.COMPANY"

  . C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

    FN_CONNECTSAPB1 $CMP $s 
   Write-Host  -ForegroundColor Green "disable SAPB1 transaction notfication SP control"

    fn_SAPB1_SP_control $ticktNum 'N' $s
  
  
    set_inacive_employee
    
   
  
  
    #Enable sapb1 TN SP control
    Write-Host -ForegroundColor Green "Enable SAPB1 transaction notfication SP control"
    $ticktNum2 =''
    fn_SAPB1_SP_control $ticktNum2 'Y' $s
    Release-Ref $cmp
  
  }