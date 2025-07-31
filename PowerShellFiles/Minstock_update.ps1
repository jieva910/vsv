# 2022.01.26
# update min stock of item 

$ticktNum      = "INC0235838"

$SourceSite    = "SR"


# load sapb1 di connection lib
. C:\Users\JIEVADM\Downloads\ps\pslib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $SourceSite


$oitem = $cmp.GetBusinessObject(4)

if ($oitem.GetByKey('0004140')){

    [int]$rows = $oitem.WhsInfo.Count                   # count all  

    for($i = 0 ; $i -lt $rows;$i++)
    {
        $oitem.WhsInfo.SetCurrentLine($i)
      if ($oitem.WhsInfo.Code -eq "SR-W"){      # only get  code equal to 
            $oitem.WhsInfo.MinimalStock =20   
           write-host $oitem.Update() $cmp.GetLastErrorDescription()
           break                                             # exit loop
      }
    }
}

