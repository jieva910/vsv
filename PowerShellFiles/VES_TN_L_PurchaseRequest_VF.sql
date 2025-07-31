USE [SAPB1_VF]
GO
/****** Object:  StoredProcedure [dbo].[VES_TN_L_PurchaseRequest]    Script Date: 6/2/2022 10:14:14 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[VES_TN_L_PurchaseRequest]

-- *** Created by: AMS, Revised by: AMS, Revision 20200416 ***

@LIndexKey AS int,
@LReply AS nvarchar(30) OUTPUT

AS

BEGIN
	SET NOCOUNT ON;

	DECLARE @PRApprovalChain  TABLE
	( approver  varchar(50) not null,
	 commoditycode  varchar(20) not null  )
		--  get PR approval chain infor		
   insert into @PRApprovalChain SELECT  Approver,  CommodityCode from  [DG-HALO01].[AURORA_VF].[dbo].[PUR_ApprovalFlowD] where StepId = 0

	SELECT @LReply = (SELECT TOP 1 * FROM (
		SELECT CASE
			WHEN 0 = 1
			THEN '8147101'	-- example scenario

		/* Start: Added by Mango Jin on 22 April 2021 - for VF (copied from LY) */
		WHEN T0.CANCELED ='N' 
			AND T0.U_VES_Status  = 'R'
			AND (T1.Price = 0
			OR ISNULL(T1.LineVendor,'') = ''
			OR ISNULL(T1.OcrCode,'') = ''
			OR ISNULL(T1.OcrCode2,'') = ''
			OR ISNULL(T1.OcrCode3,'') = ''
			OR ISNULL(T1.OcrCode4,'') = ''
			OR ISNULL(T1.OcrCode5,'') = ''
			OR ISNULL(T1.U_VES_PRP1,'')= ''
		)
		THEN '8023101'	-- PR-Validation Error/ Missing Unit Price or 5 dimensions or Supplier or purchase commodity class


		WHEN T0.CANCELED ='N' 
			AND T0.U_VES_Status  = 'R'
			--AND M.InvntItem = 'N' 
			AND T2.QryGroup19 = 'Y'  /*AND T1.ItemCode LIKE 'CAPX%'*/
			AND IsNull(T1.Project,'') = ''		
			THEN '8023102'	-- CAPEX PROJECT control				

		WHEN T0.CANCELED ='N' 
			AND T0.U_VES_Status = 'R'
			AND T2.InvntItem = 'Y'
			AND T2.EvalSystem ='S'
			AND isnull(TW.AvgPrice,0) = 0
			THEN '8023104' -- No standard cost for stock item

		WHEN T0.CANCELED ='N' 
			AND T0.U_VES_Status = 'R'
			AND T4.Status ='A'
			AND T4.U_VES_BlanketStatus in('Approved','Emailed')
			AND T1.PQTReqDate between T4.StartDate and T4.EndDate
			AND T3.LineStatus = 'O' 
			AND isnull(T1.Price,0) != T3.UnitPrice
			THEN '8023106' -- PR - Price must be same as valid blanket agreement
	  	when t0.CANCELED = 'N'
		    AND T0.U_VES_Status = 'R'
			and isnull(pra.Approver,'')='' and ISNULL(t1.linevendor,'') not like '900%'
			then '8023107' -- PR - No Approval Chain found
        
		/* End: Added by Mango Jin on 22 April 2021 - for VF (copied from LY) */

		-- *** add more scenarios below
		-- *** add more scenarios above
		END + 'ENG/Line=' + CAST(T1.VisOrder + 1 as CHAR(4))

		as 'InValid'
		FROM OPRQ T0
			INNER JOIN PRQ1 T1 ON T0.ObjType = T1.ObjType AND T0.DocEntry = T1.DocEntry
			--INNER JOIN OUSR TU ON ISNULL(T0.UserSign2, T0.UserSign) = TU.USERID

			left JOIN OITW TW ON TW.WhsCode = T1.WhsCode and T1.ItemCode = TW.ItemCode
			LEFT JOIN OITM T2 ON T1.ItemCode = t2.ItemCode
			left join OAT1 T3 ON T1.ItemCode = T3.ItemCode
			left join OOAT T4 ON T3.AgrNo = T4.AbsID and isnull(T1.LineVendor,'') = T4.BpCode and T4.BpType = 'S'
			left join ousr u on t0.UserSign = u.INTERNAL_K
		    left join @PRApprovalChain pra on pra.approver =u.USER_CODE and pra.commoditycode = t1.U_VES_PRP1

		WHERE T0.DocEntry = @LIndexKey 
	) T
	
	WHERE T.InValid IS NOT NULL	
	)
END