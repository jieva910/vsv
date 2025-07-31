# Import required assemblies for COM interop
Add-Type -TypeDefinition "using System;"

# Function to release COM objects
function Release-ComObject {
    param ([Object]$comObject)
    if ($comObject -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) | Out-Null
    }
}

# Main script
try {
    # Step 1: Initialize and connect to SAP B1
    $oCompany = New-Object -ComObject SAPbobsCOM.Company
   $oCompany.Server = "SZ-SAPSTG91"
$oCompany.SLDServer ="SZ-TSTSAPLIC92:40000"
$oCompany.CompanyDB = "SAPB1_SZ_TST"
$oCompany.DbServerType = 10
$oCompany.DbUserName ="b1if"
$oCompany.DbPassword="Vsvapp@202333"
$oCompany.UserName="montova"
$oCompany.Password="ButterSZ"

    $connectionResult = $oCompany.Connect()
    if ($connectionResult -ne 0) {
        $errorCode = 0
        $errorMsg = ""
        $oCompany.GetLastError([ref]$errorCode, [ref]$errorMsg)
        Write-Host "Connection failed: $errorMsg"
        return
    }
    Write-Host "Connected to SAP B1 successfully."

    # Step 2: Get the BatchNumberDetailsService
    $oCompanyService = $oCompany.GetCompanyService()
    $oBatchNumbersService = $oCompanyService.GetBusinessService(10000044) # ServiceTypes.BatchNumberDetailsService

    # Step 3: Define the item and batch numbers to update
    $itemCode = "LEG71097P0025@BY" # Replace with your item code
    $batchNumbers = @("BY4480344", "BY4480345", "BY4480352") # Replace with your batch numbers
    $newManufacturingDate = (Get-Date).AddDays(-30) # Example date

    # Step 4: Start a transaction
    $oCompany.StartTransaction()

    foreach ($batchNumber in $batchNumbers) {
        try {
            # Get batch parameters
            $oBatchNumberDetailParams = $oBatchNumbersService.GetDataInterface(0) # bndsBatchNumberDetailParams
            $oBatchNumberDetailParams.ItemCode = $itemCode
            $oBatchNumberDetailParams.Batch = $batchNumber

            # Retrieve the batch details
            $oBatchDetail = $oBatchNumbersService.Get($oBatchNumberDetailParams)
            if ($null -eq $oBatchDetail) {
                Write-Host "Batch $batchNumber for item $itemCode not found."
                continue
            }

            # Update batch properties
            $oBatchDetail.ManufacturingDate = $newManufacturingDate
            # Example: Update other properties as needed
            # $oBatchDetail.ExpirationDate = [DateTime]::Parse("2026-12-31")
            # $oBatchDetail.UserFields.Fields.Item("U_CustomField").Value = "NewValue"

            # Update the batch
            $oBatchNumbersService.Update($oBatchDetail)
            Write-Host "Batch $batchNumber updated successfully."
        }
        catch {
            Write-Host "Error updating batch $batchNumber : $($_.Exception.Message)"
        }
    }

    # Step 5: Commit the transaction
    $oCompany.EndTransaction(0) # BoWfTransOpt.wf_Commit
    Write-Host "All updates committed successfully."
}
catch {
    # Rollback transaction on error
    if ($oCompany.InTransaction) {
        $oCompany.EndTransaction(1) # BoWfTransOpt.wf_RollBack
    }
    Write-Host "Error: $($_.Exception.Message)"
}
<# finally {
    # Step 6: Disconnect from SAP B1 and clean up COM objects
    if ($oCompany -ne $null -and $oCompany.Connected) {
        $oCompany.Disconnect()
        Write-Host "Disconnected from SAP B1."
    }

    # Release COM objects
    Release-ComObject $oBatchDetail
    Release-ComObject $oBatchNumberDetailParams
    Release-ComObject $oBatchNumbersService
    Release-ComObject $oCompanyService
    Release-ComObject $oCompany

    # Force garbage collection to free memory
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
#>


 # Step 2: Get the BatchNumberDetailsService
    $oCompanyService = $oCompany.GetCompanyService()
    $oBatchNumbersService = $oCompanyService.GetBusinessService(10000044) # ServiceTypes.BatchNumberDetailsService

      $oBatchNumberDetail = $oBatchNumbersService.GetDataInterface(1) # bndsBatchNumberDetailParams
   
            $oBatchNumberDetail.DocEntry = 777965
             # Retrieve the batch details
            $oBatchDetail = $oBatchNumbersService.Get($oBatchNumberDetail)
                $oBatchDetail.ExpirationDate = get-date
                $oBatchNumbersService.Update($oBatchDetail)
