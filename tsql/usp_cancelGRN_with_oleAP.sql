ALTER  PROCEDURE usp_CancelPurchaseDeliveryNote
    @IndexKey INT,
    @errorDescription NVARCHAR(255) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @object INT
    DECLARE @hr INT
    DECLARE @error INT
    DECLARE @errorMsg NVARCHAR(255)
    DECLARE @companyName NVARCHAR(100)
    DECLARE @docEntry INT
    DECLARE @poObject INT
    DECLARE @POCancel INT
    DECLARE @recordset INT 
    DECLARE @retVal INT
    DECLARE @newDocEntry INT
    DECLARE @errorCode INT

    -- 初始化
    SET @hr = 0
    SET @errorDescription = ''
    
    -- 创建 SAPbobsCOM.Company 对象
    EXEC @hr = sp_OACreate 'SAPbobsCOM.Company', @object OUT
    IF @hr <> 0
    BEGIN
        EXEC @hr = sp_OAGetErrorInfo @object, @error OUT, @errorMsg OUT
        SET @errorDescription = N'步骤 2: 创建 Company 对象失败: ' + ISNULL(@errorMsg, '未知错误')
        RETURN
    END

    -- 设置连接参数
    EXEC @hr = sp_OASetProperty @object, 'Server', 'SZ-SAPSTG91'
    EXEC @hr = sp_OASetProperty @object, 'SLDServer', 'SZ-TSTSAPLIC92'
    EXEC @hr = sp_OASetProperty @object, 'DbServerType', 10 -- SQL Server 2016
    EXEC @hr = sp_OASetProperty @object, 'CompanyDB', 'SAPB1_SZ_TST'
    EXEC @hr = sp_OASetProperty @object, 'UserName', 'JIEVA'
    EXEC @hr = sp_OASetProperty @object, 'Password', 'Ves-123456'
    EXEC @hr = sp_OASetProperty @object, 'DbUserName', 'b1if'
    EXEC @hr = sp_OASetProperty @object, 'DbPassword', 'Vsvapp@202333'
    EXEC @hr = sp_OASetProperty @object, 'UseTrusted', 0

    -- 连接到公司数据库
    EXEC @hr = sp_OAMethod @object, 'Connect', @retVal OUT
    IF @retVal <> 0 AND @hr = 0 
    BEGIN
        EXEC @hr = sp_OAMethod @object, 'GetLastErrorDescription', @errorDescription OUT
        SET @errorDescription = '步骤 3: 数据库连接失败 - ' + @errorDescription
        EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
        EXEC sp_OADestroy @object
        RETURN
    END

    -- 创建 Recordset 对象用于查询
    EXEC @hr = sp_OAMethod @object,'GetBusinessObject', @recordset OUT,300
    IF @hr <> 0
    BEGIN
        EXEC sp_OAMethod @object, 'GetLastError', NULL, @errorCode OUT, @errorDescription OUT
        SET @errorDescription = '步骤 4: 创建 Recordset 对象失败 - 错误代码: ' + 
                               CAST(ISNULL(@errorCode, -1) AS NVARCHAR(10)) + 
                               ', 描述: ' + ISNULL(@errorDescription, '未知错误')
        EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
        EXEC sp_OADestroy @object
        RETURN
    END

    -- 使用 DoQuery 查询收货单状态
    DECLARE @query NVARCHAR(255)
    SET @query = 'SELECT DISTINCT DOCENTRY FROM PDN1 WHERE U_VES_BASEENTRY = ' + CAST(@IndexKey AS NVARCHAR(20))
    EXEC @hr = sp_OAMethod @recordset, 'DoQuery', NULL, @query
    IF @hr <> 0
    BEGIN
        EXEC sp_OAMethod @object, 'GetLastError', NULL, @errorCode OUT, @errorDescription OUT
        SET @errorDescription = '步骤 5: 查询失败 - 错误代码: ' + 
                               CAST(ISNULL(@errorCode, -1) AS NVARCHAR(10)) + 
                               ', 描述: ' + ISNULL(@errorDescription, '未知错误')
        EXEC sp_OADestroy @recordset
        EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
        EXEC sp_OADestroy @object
        RETURN
    END

    -- 获取查询结果
    EXEC @hr = sp_OAGetProperty @recordset, 'Fields.Item(0).Value', @docEntry OUT
    IF @hr <> 0 OR @docEntry IS NULL
    BEGIN
        SET @errorDescription = '步骤 6: 未找到收货单或查询结果为空'
        EXEC sp_OADestroy @recordset
        EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
        EXEC sp_OADestroy @object
        RETURN
    END
        
    -- 使用 GetBusinessObject 获取采购收货单
    EXEC @hr = sp_OAMethod @object, 'GetBusinessObject', @poObject OUT, 20
    IF @hr <> 0
    BEGIN
        SET @errorDescription = '步骤 7: 获取采购收货单对象失败'
        EXEC sp_OADestroy @recordset
        EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
        EXEC sp_OADestroy @object
        RETURN
    END

    -- 获取指定 DocEntry 的收货单
    EXEC @hr = sp_OAMethod @poObject, 'GetByKey', @retVal OUT, @docEntry
    IF @hr <> 0 OR @retVal = 0
    BEGIN
        EXEC sp_OAMethod @object, 'GetLastError', NULL, @errorCode OUT, @errorDescription OUT
        SET @errorDescription = '步骤 8: 加载收货单失败 - 错误代码: ' + 
                               CAST(ISNULL(@errorCode, -1) AS NVARCHAR(10)) + 
                               N', 描述: ' + ISNULL(@errorDescription, '未知错误')
        EXEC sp_OADestroy @poObject
        EXEC sp_OADestroy @recordset
        EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
        EXEC sp_OADestroy @object
        RETURN
    END

    -- 创建取消单据
    EXEC @hr = sp_OAMethod @poObject, 'CreateCancellationDocument', @POCancel OUT
    IF @hr <> 0
    BEGIN
        SET @errorDescription = '步骤 9: 创建取消单据失败'
        EXEC sp_OADestroy @poObject
        EXEC sp_OADestroy @recordset
        EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
        EXEC sp_OADestroy @object
        RETURN
    END
    
    EXEC @hr = sp_OAMethod @POCancel, 'Add', @retVal out
    IF @retVal <> 0
    BEGIN
        EXEC @hr = sp_OAMethod @object, 'GetLastErrorDescription', @errorDescription OUT
        SET @errorDescription = '步骤 10: 添加取消单据失败 - ' + @errorDescription
        EXEC sp_OADestroy @POCancel
        EXEC sp_OADestroy @poObject
        EXEC sp_OADestroy @recordset
        EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
        EXEC sp_OADestroy @object
        RETURN
    END
	else
	begin
    -- 成功完成
    SET @errorDescription = '步骤 10: 成功取消收货单，创建反向单据'
	end 
    -- 清理
    EXEC sp_OADestroy @POCancel
    EXEC sp_OADestroy @poObject
    EXEC sp_OADestroy @recordset
    EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
    EXEC sp_OADestroy @object
END
GO