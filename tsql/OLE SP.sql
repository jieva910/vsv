-- env sql server 2016 x64 on server 
-- sap b1 DI API x64 on same server  2025.04

DECLARE @object INT
DECLARE @hr INT
DECLARE @error INT
DECLARE @errorMsg NVARCHAR(255)
DECLARE @companyName NVARCHAR(100)

-- 初始化 @hr
SET @hr = 0

-- 创建 SAPbobsCOM.Company 对象
EXEC @hr = sp_OACreate 'SAPbobsCOM.Company', @object OUT
IF @hr <> 0
BEGIN
    print '无法创建 SAPbobsCOM.Company 对象'
    EXEC sp_OAGetErrorInfo @object, @error OUT, @errorMsg OUT
    print '错误: ' + ISNULL(@errorMsg, '未知错误')
    RETURN
END

-- 设置连接参数
EXEC @hr = sp_OASetProperty @object, 'Server', 'SZ-SAPSTG91' -- SQL Server 名称
EXEC @hr = sp_OASetProperty @object, 'SLDServer', 'SZ-TSTSAPLIC92' -- SLD 服务器
EXEC @hr = sp_OASetProperty @object, 'DbServerType', 10 -- SQL Server 2019
EXEC @hr = sp_OASetProperty @object, 'CompanyDB', 'SAPB1_SZ_TST' -- SAP B1 公司数据库名称
EXEC @hr = sp_OASetProperty @object, 'UserName', 'JIEVA' -- SAP 用户名
EXEC @hr = sp_OASetProperty @object, 'Password', 'Ves-123456' -- SAP 密码
EXEC @hr = sp_OASetProperty @object, 'DbUserName', 'b1if' -- SQL Server 数据库用户名
EXEC @hr = sp_OASetProperty @object, 'DbPassword', 'Vsvapp@202333' -- SQL Server 数据库密码
EXEC @hr = sp_OASetProperty @object, 'UseTrusted', 0 -- 不使用信任连接

-- 连接到公司数据库
EXEC @hr = sp_OAMethod @object, 'Connect', NULL

IF @hr <> 0
BEGIN
    EXEC sp_OAGetErrorInfo @object, @error OUT, @errorMsg OUT
    select '连接失败: ' + ISNULL(@errorMsg, '未知错误')
    EXEC sp_OADestroy @object
    RETURN
END

-- 获取公司名称
EXEC @hr = sp_OAGetProperty @object, 'CompanyName', @companyName OUT

IF @hr = 0
BEGIN
    select N'公司名称: ' + ISNULL(@companyName, '未获取到公司名称')
END
ELSE
BEGIN
    EXEC sp_OAGetErrorInfo @object, @error OUT, @errorMsg OUT
    select '获取公司名称失败: ' + ISNULL(@errorMsg, '未知错误')
END

-- 添加物料
DECLARE @itemObject INT
DECLARE @itemCode NVARCHAR(50) = 'TEST_ITEM_001' -- 物料代码
DECLARE @itemName NVARCHAR(100) = '测试物料001' -- 物料名称

-- 创建 Items 对象
EXEC @hr = sp_OAMethod @object, 'GetBusinessObject', @itemObject OUT, 4
IF @hr <> 0
BEGIN
    select '无法创建 SAPbobsCOM.Items 对象'
    EXEC sp_OAGetErrorInfo @object, @error OUT, @errorMsg OUT
    select '错误: ' + ISNULL(@errorMsg, '未知错误')
    EXEC sp_OADestroy @object
    RETURN
END

-- 设置物料属性
EXEC @hr = sp_OASetProperty @itemObject, 'ItemCode', @itemCode
EXEC @hr = sp_OASetProperty @itemObject, 'ItemName', @itemName
EXEC @hr = sp_OASetProperty @itemObject, 'ItmsGrpCod', 143 -- 物料组代码

-- 添加物料
EXEC @hr = sp_OAMethod @itemObject, 'Add', NULL
IF @hr <> 0
BEGIN
    EXEC sp_OAGetErrorInfo @itemObject, @error OUT, @errorMsg OUT
    select '添加物料失败: ' + ISNULL(@errorMsg, '未知错误')
END
ELSE
BEGIN
    select '物料 ' + @itemCode + ' 添加成功'
END

-- 清理物料对象
EXEC sp_OADestroy @itemObject

-- 断开连接并销毁 Company 对象
EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
EXEC sp_OADestroy @object