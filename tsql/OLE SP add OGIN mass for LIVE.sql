
use [sz-sapstg92]
go 

 -- execute in MSDB

ALTER PROCEDURE [dbo].[usp_ProcessConsignment]
@Result int output,
@errorDescription nvarchar(255) output
with Encryption
AS 
BEGIN
    SET NOCOUNT ON;
    DECLARE 
        @hr INT,
        @cmp INT,
		@logMessage NVARCHAR(4000)='',
		 @i int,
		 @j int ,
		 @batchNumbers int ;
		
    -- 执行查询 LINK SERVER TO LIVE 
 SELECT *,format(BatchManufacturingDate,'yyyy-MM-dd') BatchManufacturingDate2 into #TempData FROM [SZ-SAP01].[sapb1-sz].dbo.VES_AI_ConsignmentService ;
   if not exists (SELECT 1 FROM #TempData)
    BEGIN
	 set @Result = -1 
        set @errorDescription =  'No records found';
        RETURN;
    END
    -- 创建SAP Company对象
    EXEC @hr = sp_OACreate 'SAPbobsCOM.Company', @cmp OUT;
    IF @hr <> 0 BEGIN
        EXEC sp_OAGetErrorInfo @cmp, NULL, @logMessage OUT;
		set @errorDescription =  'failed to create SAP COM object: ' + @logMessage;
		set @Result = -1
      --  EXEC LogMessage @logMessage
        RETURN;
    END

	    EXEC @hr = sp_OASetProperty @cmp, 'Server', 'sz-sap01'
    EXEC @hr = sp_OASetProperty @cmp, 'SLDServer','sz-saplic92'
    EXEC @hr = sp_OASetProperty @cmp, 'DbServerType',8
    EXEC @hr = sp_OASetProperty @cmp, 'CompanyDB', 'sapb1-SZ'
    EXEC @hr = sp_OASetProperty @cmp, 'UserName', 'Montova'
    EXEC @hr = sp_OASetProperty @cmp, 'Password', 'ButterSZ'
    EXEC @hr = sp_OASetProperty @cmp, 'DbUserName', 'Butterfly'
    EXEC @hr = sp_OASetProperty @cmp, 'DbPassword', 'buTterF1y'
    EXEC @hr = sp_OASetProperty @cmp, 'UseTrusted', 0

    -- 连接SAP
    EXEC @hr = sp_OAMethod @cmp, 'Connect', NULL;
    IF @hr <> 0 BEGIN
        EXEC sp_OAGetErrorInfo @cmp, NULL, @logMessage OUT;
		set @errorDescription = 'failed to connect to SAPB1: ' + @logMessage;
		set @Result = -1 
       -- EXEC LogMessage  @logMessage,NULL
        RETURN;
    END


    -- 主分组处理逻辑
    DECLARE @DocDate DATETIME, @TaxDate DATETIME, @Reference2 NVARCHAR(255), @sapError nvarchar(255),
            @Comments NVARCHAR(255), @PaymentGroupCode NVARCHAR(50), @DocEntry INT;

    DECLARE headerCursor CURSOR LOCAL FAST_FORWARD FOR 
    SELECT DISTINCT 
        Comments,
        Reference2,
        DocDate,
        TaxDate,
        PaymentGroupCode,
        DocEntry
    FROM #TempData;

    OPEN headerCursor;
    FETCH NEXT FROM headerCursor INTO 
        @Comments, @Reference2, @DocDate, @TaxDate, @PaymentGroupCode, @DocEntry;

	IF DATEDIFF(MM,@DocDate,GETDATE())<>0
	SET @DocDate = FORMAT(GETDATE(),'yyyy-MM-dd')
    WHILE @@FETCH_STATUS = 0
    BEGIN
        DECLARE @oOIGN INT, @rtCode INT=0, @oDln INT ;
        
        BEGIN TRY
            -- 开始事务
            EXEC @hr = sp_OAMethod @cmp, 'StartTransaction', NULL;
            -- 创建OIGN对象
            EXEC @hr = sp_OAMethod @cmp, 'GetBusinessObject', @oOIGN OUT, 59;
            -- 设置Header属性
            EXEC @hr = sp_OASetProperty @oOIGN, 'DocDate', @DocDate;
            EXEC @hr = sp_OASetProperty @oOIGN, 'TaxDate', @TaxDate;
            EXEC @hr = sp_OASetProperty @oOIGN, 'Reference2', @Reference2;
            EXEC @hr = sp_OASetProperty @oOIGN, 'Comments', @Comments;
            EXEC @hr = sp_OASetProperty @oOIGN, 'PaymentGroupCode', @PaymentGroupCode;

            -- 处理行项目
            DECLARE @ItemCode NVARCHAR(50), @Quantity DECIMAL(19,6), @UnitPrice DECIMAL(19,6),
                    @Currency NVARCHAR(3), @AccountCode NVARCHAR(50), @CostingCode NVARCHAR(50),
                    @CostingCode2 NVARCHAR(50), @CostingCode3 NVARCHAR(50), @CostingCode4 NVARCHAR(50),
                    @CostingCode5 NVARCHAR(50), @ManufSource NVARCHAR(50), @U_FinUse NVARCHAR(50), @lines int ;

            DECLARE lineCursor CURSOR LOCAL FAST_FORWARD FOR 
           select ItemCode,  sum(Quantity) Quantity, UnitPrice, Currency, AccountCode,
                CostingCode, CostingCode2, CostingCode3, CostingCode4, CostingCode5,
                ManufSource, U_FinUse from (
  SELECT  distinct
                ItemCode,  Quantity, UnitPrice, Currency, AccountCode,
                CostingCode, CostingCode2, CostingCode3, CostingCode4, CostingCode5,
                ManufSource, U_FinUse,LineNum
            FROM  #TempData 
            WHERE DocEntry =@DocEntry ) as DistinctRows  group by ItemCode ,UnitPrice, Currency, AccountCode,
                CostingCode, CostingCode2, CostingCode3, CostingCode4, CostingCode5,
                ManufSource, U_FinUse ;


            OPEN lineCursor;
            FETCH NEXT FROM lineCursor INTO @ItemCode, @Quantity, @UnitPrice, @Currency, @AccountCode,
                @CostingCode, @CostingCode2, @CostingCode3, @CostingCode4, @CostingCode5, @ManufSource, @U_FinUse;
             set @i = 0
            WHILE @@FETCH_STATUS = 0
            BEGIN
			  EXEC @hr = sp_OAMethod @oOIGN, 'Lines', @lines OUT ;
             
                EXEC @hr = sp_OASetProperty @lines, 'ItemCode', @ItemCode;
                EXEC @hr = sp_OASetProperty @lines, 'Quantity', @Quantity;
                EXEC @hr = sp_OASetProperty @lines, 'UnitPrice', @UnitPrice;
                EXEC @hr = sp_OASetProperty @lines, 'Currency', @Currency;
                EXEC @hr = sp_OASetProperty @lines, 'WarehouseCode', 'CT';
                EXEC @hr = sp_OASetProperty @lines, 'AccountCode', @AccountCode;
                EXEC @hr = sp_OASetProperty @lines, 'CostingCode', @CostingCode;
                EXEC @hr = sp_OASetProperty @lines, 'CostingCode2', @CostingCode2;
                EXEC @hr = sp_OASetProperty @lines, 'CostingCode3', @CostingCode3;
                EXEC @hr = sp_OASetProperty @lines, 'CostingCode4', @CostingCode4;
                EXEC @hr = sp_OASetProperty @lines, 'CostingCode5', @CostingCode5;
                
                -- 用户字段
               DECLARE @userFields INT;
                EXEC @hr = sp_OAMethod @lines, 'UserFields', @userFields OUT;
                EXEC @hr = sp_OASetProperty @userFields, 'Fields.Item("U_manufsource").Value', @ManufSource;
                EXEC @hr = sp_OASetProperty @userFields, 'Fields.Item("U_FinUse").Value', @U_FinUse;
            

				-- set @logMessage = concat('lines:',@itemcode,' ',@Quantity,' ',@i)
				--   EXEC LogMessage  @logMessage,@DocEntry ;

                -- 处理批次
                DECLARE @BatchNumber NVARCHAR(36), @BatchQuantity DECIMAL(19,6), 
                        @Batch_U_FinUse NVARCHAR(50), @BatchManufacturingDate2 DATE,
						@U_BoxDepth DECIMAL(19,6),	@U_BoxHeight DECIMAL(19,6),	@U_BoxWidth DECIMAL(19,6),	@U_Net DECIMAL(19,6),	@U_Gross DECIMAL(19,6),@U_BoxType nvarchar(50) ;

                DECLARE batchCursor CURSOR LOCAL FAST_FORWARD FOR 
                SELECT distinct
                    CASE WHEN LEN(BatchNumber) > 36 THEN cast(ReferencedDocNumber as nvarchar(30)) + CAST(LineNum AS NVARCHAR) + '_' + U_FinUse
                         ELSE BatchNumber END AS BatchNumber,
                    BatchQuantity,
                    U_FinUse,
				   BatchManufacturingDate2 ,
				   U_BoxDepth,	U_BoxHeight,	U_BoxWidth,	U_Net,	U_Gross,U_BoxType
                FROM #TempData 
                WHERE DocEntry = @DocEntry 
                  AND ItemCode = @ItemCode;

                OPEN batchCursor;
                FETCH NEXT FROM batchCursor INTO @BatchNumber, @BatchQuantity, @Batch_U_FinUse, @BatchManufacturingDate2, @U_BoxDepth,@U_BoxHeight,@U_BoxWidth,@U_Net,	@U_Gross,@U_BoxType;

				set @j = 0 
                WHILE @@FETCH_STATUS = 0
                BEGIN
                   EXEC @hr = sp_OAMethod @lines, 'BatchNumbers', @batchNumbers OUT;
                    EXEC @hr = sp_OASetProperty @batchNumbers, 'BatchNumber', @BatchNumber;
                    EXEC @hr = sp_OASetProperty @batchNumbers, 'AddmisionDate', @BatchManufacturingDate2;
                    EXEC @hr = sp_OASetProperty @batchNumbers, 'ManufacturingDate', @BatchManufacturingDate2;
                    EXEC @hr = sp_OASetProperty @batchNumbers, 'Quantity', @BatchQuantity;
                    
                   DECLARE @batchFields INT;
                    EXEC @hr = sp_OAMethod @batchNumbers, 'UserFields', @batchFields OUT;
                    EXEC @hr = sp_OASetProperty @batchFields,'Fields.Item("U_FinUse").Value', @Batch_U_FinUse;
                  EXEC @hr = sp_OASetProperty @batchFields,'Fields.Item("U_BoxDepth").Value', @U_BoxDepth;
				   EXEC @hr = sp_OASetProperty @batchFields,'Fields.Item("U_BoxHeight").Value', @U_BoxHeight;
				    EXEC @hr = sp_OASetProperty @batchFields,'Fields.Item("U_BoxWidth").Value', @U_BoxWidth;
					 EXEC @hr = sp_OASetProperty @batchFields,'Fields.Item("U_Net").Value', @U_Net;
					  EXEC @hr = sp_OASetProperty @batchFields,'Fields.Item("U_Gross").Value', @U_Gross;
					   EXEC @hr = sp_OASetProperty @batchFields,'Fields.Item("U_BoxType").Value', @U_BoxType;
					    EXEC @hr = sp_OASetProperty @batchFields,'Fields.Item("U_manufsource").Value', @ManufSource;
					-- set @logMessage = concat('batchs:',@itemcode,' ',@BatchNumber,' ',@BatchQuantity,' ',@j)
					--  EXEC LogMessage  @logMessage,@DocEntry ;

                    FETCH NEXT FROM batchCursor INTO @BatchNumber, @BatchQuantity, @Batch_U_FinUse, @BatchManufacturingDate2, @U_BoxDepth,@U_BoxHeight,@U_BoxWidth,@U_Net,	@U_Gross,@U_BoxType;
					
					  EXEC @hr = sp_OAMethod @batchNumbers, 'SetCurrentLine',NULL,@j

					 EXEC @hr = sp_OAMethod @batchNumbers,  'Add',NULL;
					 set @j = @j + 1 ; 
					
                END

                CLOSE batchCursor;
                DEALLOCATE batchCursor;

                FETCH NEXT FROM lineCursor INTO @ItemCode, @Quantity, @UnitPrice, @Currency, @AccountCode,
                    @CostingCode, @CostingCode2, @CostingCode3, @CostingCode4, @CostingCode5, @ManufSource, @U_FinUse;

                 	EXEC @hr = sp_OAMethod @lines, 'SetCurrentLine',NULL,@i 
                    EXEC @hr = sp_OAMethod @lines,  'Add',NULL;
                  set @i= @i+1 ;
				 
            END

            CLOSE lineCursor;
            DEALLOCATE lineCursor;

            -- 提交OIGN
            EXEC @hr = sp_OAMethod @oOIGN, 'Add', @rtCode OUT;
			IF @rtCode <> 0 
				BEGIN
					-- 获取 SAP 底层错误描述
					EXEC @hr = sp_OAMethod @cmp, 'GetLastErrorDescription', @sapError OUT;
    
					-- 抛出包含 SAP 错误信息的异常
					DECLARE @errorMsg NVARCHAR(4000) = 
						CONCAT('add Goods receipt failure | SAP error code: ', @rtCode, 
							   ' | msg: ', @sapError
							  );
					--exec logMessage @errorMsg,@DocEntry;
					set @errorDescription = @errorMsg;
					set @Result = -1 
					RAISERROR(@errorDescription, 16, 1);
				END
		
				-- 关闭ODLN
				EXEC @hr = sp_OAMethod @cmp, 'GetBusinessObject', @oDln OUT, 15;
				EXEC @hr = sp_OAMethod @oDln, 'GetByKey',@rtCode OUT,  @DocEntry;
				EXEC @hr = sp_OAMethod @oDln, 'Close', @rtCode OUT;

				IF @rtCode <> 0 
				BEGIN
					EXEC @hr = sp_OAMethod @cmp, 'GetLastErrorDescription', @sapError OUT;
					set @errorMsg  = 
						CONCAT('close failure | SAP error code: ', @rtCode, 
							   ' | msg: ', @sapError
							 );
					set @errorDescription = @errorMsg; 
					set @Result = -1 
					RAISERROR(@errorDescription, 16, 1);
				END
	             -- 提交事务
        EXEC @hr = sp_OAMethod @cmp, 'EndTransaction', NULL, 0;
        set @errorDescription= 'success for delivery note Entry:'+ cast(@DocEntry as varchar);
          set @Result = 0 
        END TRY
        BEGIN CATCH
				-- 回滚事务
			EXEC @hr = sp_OAMethod @cmp, 'EndTransaction', NULL, 1;
    
			-- 记录详细错误日志
			set @errorMsg  = 
				CONCAT('Error: ', ERROR_MESSAGE(), 
					   ' | current server time: ', GETDATE());
    
			--EXEC LogMessage @errorMsg, @DocEntry;
			set @errorDescription =@errorMsg
			set @Result = -1 
			
        END CATCH

        FETCH NEXT FROM headerCursor INTO 
           @Comments, @Reference2, @DocDate, @TaxDate, @PaymentGroupCode, @DocEntry;
    END

    CLOSE headerCursor;
    DEALLOCATE headerCursor;

    -- 清理对象
    EXEC sp_OADestroy @oOIGN;
    EXEC sp_OADestroy @oDln;
	if @cmp is NOT NULL
	BEGIN
      EXEC @hr = sp_OAMethod @cmp, 'Disconnect', NULL
      EXEC sp_OADestroy @cmp
	end 
   
END
