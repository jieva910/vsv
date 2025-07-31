$cmp = New-Object -COMobject "SAPbobsCOM.Company"
 
# Release COM object 
function Release-Ref ($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()}
 
function UI_DI_Conn
{
 
    

    # Connect to SBO via UI API
        Function SetApplication {
              $SboGuiApi = New-Object -comobject "SAPbouiCOM.SboGuiApi"
              $sConnectionString =  "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

              $SboGuiApi.Connect($sConnectionString)
              $SboGuiApi.GetApplication()
          
        }

    #Connect with connection string
       $SBO_Application = SetApplication

        function SetConnectionContext {
    
             $sCookie = $cmp.GetContextCookie()
             $sConnectionContext = $SBO_Application.Company.GetConnectionContext($sCookie)
             If ($cmp.Connected ){$cmp.Disconnect()}
             return $cmp.SetSboLoginContext($sConnectionContext)
        }

    # Connect to SBO via DI API

    Function ConnectTcmp {
       Return $cmp.Connect()
    }

    # connect to DI 
   
        if (SetConnectionContext -ne 0 ) {$SBO_Application.MessageBox("Failed setting a connection to DI API");break}
        if (ConnectTcmp -ne 0 ) {$SBO_Application.MessageBox("Failed connecting to the company's Data Base") ; break}
    Write-Host -ForegroundColor Cyan "DI Connected To: " $cmp.LicenseServer $cmp.CompanyName
}

UI_DI_Conn


	$oApprovalRequestsService = $cmp.GetCompanyService().GetBusinessService(122)    # ApprovalRequestsService=122
	 $oApprovalRequestsParams =   $oApprovalRequestsService.GetDataInterface(1)       # arsApprovalRequestsParams = 1 
	 $oApprovalRequest =$oApprovalRequestsService.GetDataInterface(0) # arsApprovalRequest = 0
	 $oApprovalRequestParams = $oApprovalRequestsService.GetDataInterface(2)

	 # Get request list 
	$oApprovalRequestsParams = $oApprovalRequestsService.GetAllApprovalRequestsList() 
	$oApprovalRequestParams = $oApprovalRequestsParams.Item($oApprovalRequestsParams.Count - 1) 

	# Approve request  
	$oApprovalRequest = $oApprovalRequestsService.GetApprovalRequest($oApprovalRequestParams) 
	$oApprovalRequestDecision = $oApprovalRequest.ApprovalRequestDecisions.Add() 
	$oApprovalRequestDecision.Remarks = " not  Approved by mango jin  in ps" 
	$oApprovalRequestDecision.Status  = 2

	 
	# Incase we want to approve with another user, uncomment the following 2 lines ,B1User must be an Authorizer in Approval Stage
	$oApprovalRequestDecision.ApproverUserName = 'jinman'                                   
	$oApprovalRequestDecision.ApproverPassword = 'Ves-123456' 


	Try 
		{ $oApprovalRequestsService.UpdateRequest($oApprovalRequest) }
	Catch { $ex = $_.Exception }
	   finally { $ex.Message }




# =========== 以下是批量操作============================================

	$sql = "-- 找出approval stage 关联的pending 或者approved with not generated 的documents 
	with Approvalrequest
	as 
	(
	SELECT 
	 t0.WddCode,
	t0.WtmCode,
	t0.DocEntry,
	t0.DocDate,
	t0.Status,
	t0.Remarks,
	t0.IsDraft,
	t1.wddstatus 'Status on Draft Doc.' ,
	t.status 'approvedbywho',
	t5.user_code,
	t5.U_name,
	t3.name 'ApprovalTemplate',
	row_number() over (partition by t0.wddcode order by t.status desc ) rid
	FROM [dbo].[OWDD]  T0 
	inner join wdd1 t on t0.wddcode = t.wddcode
	 left join odrf t1 on t0.docentry = t1.docentry 
	left join owst t2 on t0.CurrStep = t2.wstcode
	inner join owtm t3 on t0.WtmCode = t3.WtmCode
	left join ousr t5 on t.userid = t5.userid 
	where t0.isdraft = 'Y'        -- 草稿状态
	and t0.status <>'N'            -- 非reject 状态
	and t.status in ('w', 'Y')   --W 表示PENDING,Y 表示approved
	and t2.name = 'ARInvoices_GB'    --appvoval stage name
	)

	select * from Approvalrequest where rid = 1 "

	$oRs = $cmp.GetBusinessObject(300)
	$oApprovalRequestsService = $cmp.GetCompanyService().GetBusinessService(122)    # ApprovalRequestsService=122

	$ors.DoQuery($sql)

	if (!$oRs.EoF)
	{
	  [xml]$Nodes = $ors.GetAsXML()
	  
	  $Rows = $Nodes.SelectNodes("//row")

	  foreach($r in $Rows)
	  {
	   # Approve request 
		   $oApprovalRequest =$oApprovalRequestsService.GetDataInterface(0)                 # arsApprovalRequest = 0
		   $oApprovalRequestParams = $oApprovalRequestsService.GetDataInterface(2) 
		   $oApprovalRequestParams.Code=$r.WddCode
			$oApprovalRequest = $oApprovalRequestsService.GetApprovalRequest($oApprovalRequestParams) 
			$oApprovalRequestDecision = $oApprovalRequest.ApprovalRequestDecisions.Add() 
			$oApprovalRequestDecision.Remarks = " 2ND mass rejection  in ps 2021.11.19 for SR TEST " 
			$oApprovalRequestDecision.Status  = 2
			# In case we want to approve with another user, uncomment the following 2 lines ,B1User must be an Authorizer in Approval Stage
			$oApprovalRequestDecision.ApproverUserName = $r.user_code                                  
			$oApprovalRequestDecision.ApproverPassword = 'Ves-123456' 


			Try 
				{ $oApprovalRequestsService.UpdateRequest($oApprovalRequest) }
			Catch { $ex = $_.Exception }
			finally { Write-Host $r.WddCode  $ex.Message }
			
			 Release-Ref $oApprovalRequest
	  }
	}
