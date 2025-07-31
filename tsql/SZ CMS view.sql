	SELECT 
			'C' as ReportType, 
			cast(p.U_VES_SCNum as varchar(10)) as YANumber, 
			D.Manager + '-' +  cast(A.DocEntry as NVARCHAR(20)) as YBUniqueId, 
			T0.Project as YBNumber, 		
			cast(left(A.ImportEnt,10) as smallint) as AgreementVersion, 
			A.Indicator as VersionStatus, 
			--V.CardCode as CustomerCode,  --bill to code  
			(Select CardCode from OQUT where DocNum= t0.AgrDocNum) as CustomerCode, -- shipto code 
			CONVERT(CHAR(6), T0.ConsPeriod, 112) as YearMonthConsumed as CustomerCode,
			CONVERT(CHAR(6), T0.DocDate, 112) as YearMonthFiscal, 
			CAST(YEAR(T0.ConsPeriod) *1000 + DATEPART(dayofyear, T0.ConsPeriod) AS VARCHAR(60)) as YearDayConsumed, 
			CAST(YEAR(T0.DocDate) *1000 + DATEPART(dayofyear, EoMonth(T0.DocDate)) AS VARCHAR(60)) as YearDayFiscal, 
			T0.ItemCode as ItemCode, 
			CAST(T0.ItemType AS smallint) as PatternType, 		
			CAST(dbo.[fn_VES_CMS_GetPatternTypeDesc] (T0.ItemType, 0 ) AS varchar(10)) as PatternTypeDesc, 
			CAST(LEFT(upper(T0.YBUoM), 2) AS nvarchar(2)) as YBUOM, 
			V.DocCur as Currency, 
			' ' as LocalControl1, 
			' ' as LocalControl2, 
			CAST(LEFT(T0.ManufOrigin, 30) AS nvarchar(35)) as ManufOrigin, 
			D.Manager as InvoiceCompany, 
			'CT' as Warehouse, 
			'Won' as AgreementStatus, 
			'Approved' as ConsumptionStatus, 
			CAST(0 AS smallint) as BillingType,  
			UPPER(LEFT(T0.ProductUOM,2)) as ProductUOM, 
			CAST(T0.UnitCOGS AS float) as UnitCOGS, 
			CAST(T0.UnitSTDCost AS float) as StdCost, 
			CAST(T0.Quantity AS float) as ConsumedQty, 
			CAST(iif(T0.Quantity=0, 0, T0.InvoiceRevenue) AS float) as ConsumedRevenue, 		
			CAST(iif(T0.Quantity=0 and T0.ContractRevenue <>0 , iif(T0.ContractRevenue>0, 1.0, -1.0), T0.Quantity) AS float) AS ContractQty, 
			CAST(iif(T0.Quantity=0 and T0.ContractRevenue <>0 , iif(T0.ContractRevenue>0, 1.0, -1.0), T0.Quantity) AS float) AS ContractCostQty,
			CAST(T0.ContractRevenue AS float) as ContractRevenue, 
			CAST(T0.ContractYBQty AS float) as ContractYBQty, 
			CAST(T0.ContractYBRevenue AS float) as ContractYBRevenue, 
			CAST(iif(T0.Quantity=0 and T0.InvoiceRevenue <>0 , iif(T0.InvoiceRevenue>0, 1.0, -1.0), T0.Quantity) AS float) as InvoiceQty, 
			CAST(iif(T0.Quantity=0 and T0.InvoiceRevenue <>0 , iif(T0.InvoiceRevenue>0, 1.0, -1.0), T0.Quantity) AS float) as InvoiceCostQty, 
			CAST(T0.InvoiceRevenue AS float) as InvoiceRevenue, 
			CAST(T0.InvoiceYBQty AS float) as InvoiceYBQty, 
			CAST(T0.InvoiceYBRevenue AS float) as InvoiceYBRevenue, 
			CAST(0.0 AS float) as OpeningQty, 
			CAST(0.0 AS float) as ReceivedQty, 
			CAST(0.0 AS float) as InTransitQty, 
			CAST(0.0 AS float) as AdjustedQty, 
			CAST(0.0 AS float) as ClosingQty, 
			CAST(CASE WHEN ISNULL(V.U_VES_CMS_ProdType, ISNULL(A.U_VES_CMS_ProdType,'')) = 'I' THEN V.U_VES_CMS_ProdVal END as float) as IronProduction, 
			CAST(CASE WHEN ISNULL(V.U_VES_CMS_ProdType, ISNULL(A.U_VES_CMS_ProdType,'')) = 'L' THEN V.U_VES_CMS_ProdVal END as float) as LiquidSteelProduction, 
			CAST(CASE WHEN ISNULL(V.U_VES_CMS_ProdType, ISNULL(A.U_VES_CMS_ProdType,'')) = 'S' THEN V.U_VES_CMS_ProdVal END as float) as SolidSteelProduction, 
			CAST(V.U_VES_CMS_Heats AS float) as NumberHeats, 
			CAST(V.U_VES_CMS_Tundish AS float) as NumberTundish, 
			CAST(V.U_VES_CMS_EAFLife AS float) as EAFLife, 
			CAST(V.U_VES_CMS_EAFDLife AS float) as EafDeltaLife, 
			CAST(V.U_VES_CMS_TorLife AS float) as TorpedoLife, 
			Case when ISNULL(V.U_VES_CMS_Ladles,0) > 0 then cast(isnull(V.U_VES_CMS_Heats,0) / A.U_VES_CMS_Ladles as float) Else 0 END AS LadleLife,
			CAST(V.U_VES_CMS_TapLife AS float) as TapholeLife, 
			CAST(V.U_VES_CMS_RunLife AS float) as RunnersLife, 
			CAST(V.U_VES_CMS_BOFLife AS float) as BOFLife, 
			CAST(V.U_VES_CMS_HCount AS float) as HeadCount, 
			CAST(V.U_VES_CMS_Strand AS float) as Strand, 
			CAST(V.U_VES_CMS_HeatSize AS float) as HeatSize, 
			CAST(V.U_VES_CMS_Porous AS float) as PorousPlugLadle, 
			CAST(V.U_VES_CMS_Gates AS float) as GatesLadle, 
			CAST(LEFT(ISNULL(V.NumAtCard,''),50) AS nvarchar(100)) as CustomerPO, 
			T0.DocNum as VesInvoiceNbr, 
			P.U_VES_CMS_GoLive as GoLiveDate, 
			CAST(LEFT(T0.LNProductApp, 10) AS nvarchar(50)) as LNProductApp
		
	FROM [dbo].[tbl_VES_CMS_INVL] T0 WITH (NOLOCK)
		LEFT OUTER JOIN OPRJ p WITH (NOLOCK) ON T0.Project = p.PrjCode
		LEFT OUTER JOIN OQUT A WITH (NOLOCK) ON T0.AgrDocNum = A.DocNum
		LEFT OUTER JOIN OINV V WITH (NOLOCK) ON T0.DocEntry = V.DocEntry
		LEFT OUTER JOIN OADM D WITH (NOLOCK) ON 1=1
		INNER JOIN dbo.vw_VES_CMS_Version_CutOff R ON R.IsCurrent ='Y' -- DO NOT REMOVE THIS LINE !!!!!  ---
	WHERE 
		T0.DocDate BETWEEN R.Period_FiscalDateFrom AND R.Period_FiscalDateTo  -- DO NOT REMOVE THIS LINE !!!!!  ---
		AND T0.DocDate BETWEEN R.Period_FiscalDateFrom AND R.Period_FiscalDateTo -- DO NOT REMOVE THIS LINE !!!!!  ---
		AND CONVERT(CHAR(6), T0.DocDate, 112) >= R.Period_BIUpload_FromPeriod -- DO NOT REMOVE THIS LINE !!!!!  ---


	union all 

	------------------------------------------------------------
	/*** MJ: 2023-08-03   
		This is the archive data before 2023-08-01 which download the copy from BI table to local SAP  
		The old query uploaded WRONG value for YBUOM into BI, I manually corrected YBUOM for fiscal period between 202301 and 202307 in both local below table and BI table.
	**/

		SELECT [ReportType]
			  ,[YANumber],[YBUniqueId],[YBNumber]
			  ,[AgreementVersion],[VersionStatus]
			  ,[CustomerCode]
			  ,[YearMonthConsumed],[YearMonthFiscal],[YearDayConsumed],[YearDayFiscal]
			  ,T0.[Pattern] as ItemCode
			  ,[PatternType],[PatternTypeDesc]
			  ,[YBUOM],[Currency]
			  ,[LocalControl1],[LocalControl2],[ManufOrigin],[InvoiceCompany],[Warehouse]
			  ,[AgreementStatus],[ConsumptionStatus]
			  ,[BillingType],[ProductUOM]
			  ,[UnitCOGS],[StdCost]
			  ,[ConsumedQty],[ConsumedRevenue]
			  ,[ContractQty],[ContractCostQty],[ContractRevenue],[ContractYBQty],[ContractYBRevenue]
			  ,[InvoiceQty],[InvoiceCostQty],[InvoiceRevenue],[InvoiceYBQty],[InvoiceYBRevenue]
			  ,[OpeningQty],[ReceivedQty],[InTransitQty],[AdjustedQty],[ClosingQty]
			  ,[IronProduction],[LiquidSteelProduction],[SolidSteelProduction],[NumberHeats],[NumberTundish]
			  ,[EAFLife],[EafDeltaLife],[TorpedoLife]
			  ,[LadleLife],[TapholeLife],[RunnersLife],[BOFLife]
			  ,[HeadCount],[Strand],[HeatSize],[PorousPlugLadle],[GatesLadle]
			  ,[CustomerPO]
			  ,[VesInvoiceNbr]
			  ,[GoLiveDate]
			  ,[LNProductApp]
		  FROM [dbo].[tbl_VES_CMS_ConsumeBI_Combined_Archive_20230731_DONOT_DELETE] T0
			, dbo.vw_VES_CMS_Version_CutOff R
		WHERE 
			R.IsCurrent ='N'  -- DO NOT REMOVE THIS LINE !!!!!  ---

			-- DO NOT REMOVE THIS LINE !!!!!  ---
			AND T0.YearMonthFiscal BETWEEN convert(varchar(6), R.Period_FiscalDateFrom, 112) AND CONVERT(VARCHAR(6), R.Period_FiscalDateTo,112)  -- DO NOT REMOVE THIS LINE !!!!!  ---
			AND T0.YearMonthFiscal >= R.Period_BIUpload_FromPeriod -- DO NOT REMOVE THIS LINE !!!!!  ---
