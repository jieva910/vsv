-- env SQL Server 2016 x64 on server 
-- SAP B1 DI API x64 on same server  2025.04

DECLARE @object INT
DECLARE @hr INT
DECLARE @error INT
DECLARE @errorMsg NVARCHAR(255)
DECLARE @companyName NVARCHAR(100)
DECLARE @docEntry INT
DECLARE @poObject INT, @POCancel INT
DECLARE @retVal INT
DECLARE @newDocEntry INT
DECLARE @errorCode INT
DECLARE @errorDescription NVARCHAR(255)

-- 设置要取消的收货单 DocEntry
SET @docEntry = 269667

-- 初始化 @hr
SET @hr = 0

-- 创建 SAPbobsCOM.Company 对象
EXEC @hr = sp_OACreate 'SAPbobsCOM.Company', @object OUT
IF @hr <> 0
BEGIN
    EXEC @hr = sp_OAGetErrorInfo @object, @error OUT, @errorMsg OUT
    SELECT N'创建 Company 对象失败: ' + ISNULL(@errorMsg, '未知错误')
    RETURN
END

-- 设置连接参数
EXEC @hr = sp_OASetProperty @object, 'Server', 'SZ-SAPSTG91'
EXEC @hr = sp_OASetProperty @object, 'SLDServer', 'SZ-TSTSAPLIC92'
EXEC @hr = sp_OASetProperty @object, 'DbServerType',10 -- SQL Server 2016
EXEC @hr = sp_OASetProperty @object, 'CompanyDB', 'SAPB1_SZ_TST'
EXEC @hr = sp_OASetProperty @object, 'UserName', 'JIEVA'
EXEC @hr = sp_OASetProperty @object, 'Password', 'Ves-123456'
EXEC @hr = sp_OASetProperty @object, 'DbUserName', 'b1if'
EXEC @hr = sp_OASetProperty @object, 'DbPassword', 'Vsvapp@202333'
EXEC @hr = sp_OASetProperty @object, 'UseTrusted', 0

-- 连接到公司数据库
EXEC @hr = sp_OAMethod @object, 'Connect', @retVal OUT

IF @retVal <> 0 AND @hr =  0 
BEGIN
    EXEC  @hr = sp_OAMethod @object, 'GetLastErrorDescription', @errorDescription OUT
    SELECT @errorDescription 'CONNFAIL'
    GOTO Cleanup
END

-- 获取公司名称
EXEC @hr = sp_OAGetProperty @object, 'CompanyName', @companyName OUT
IF @hr = 0
BEGIN
    SELECT N'公司名称: ' + ISNULL(@companyName, '未获取到公司名称')
END

-- 使用 GetBusinessObject 获取采购收货单 (对象类型 20)
EXEC @hr = sp_OAMethod @object, 'GetBusinessObject', @poObject OUT, 20 -- bo_PurchaseDeliveryNotes = 20


-- 获取指定 DocEntry 的收货单
EXEC @hr = sp_OAMethod @poObject, 'GetByKey', @retVal OUT, @docEntry
IF @hr <> 0 OR @retVal = 0
BEGIN
    EXEC sp_OAMethod @object, 'GetLastError', NULL, @errorCode OUT, @errorDescription OUT
    SELECT N'加载收货单失败 - 错误代码: ' + CAST(ISNULL(@errorCode, -1) AS NVARCHAR(10)) + 
           N', 描述: ' + ISNULL(@errorDescription, '未知错误')
    GOTO CleanupPO
END

-- 创建取消单据
EXEC @hr = sp_OAMethod @poObject, 'CreateCancellationDocument', @POCancel OUT
 EXEC @hr = sp_OAMethod @POCancel, 'Add', Null

  IF @hr = 0
    BEGIN
        SELECT N'成功取消收货单，创建反向单据 DocEntry: ' 
    END
    ELSE
    BEGIN
        EXEC @hr = sp_OAMethod @object, 'GetLastErrorDescription', @errorDescription OUT
		 SELECT  @errorDescription
    END


   

-- 清理采购收货单对象
CleanupPO:
IF @poObject IS NOT NULL
BEGIN
    EXEC sp_OADestroy @poObject
END

-- 清理公司对象
Cleanup:
IF @object IS NOT NULL
BEGIN
    EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
    EXEC sp_OADestroy @object
END