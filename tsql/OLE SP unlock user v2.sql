
alter PROCEDURE [dbo].[usp_UnlockUser2]
    @USER VARCHAR(10),
	@CompanyDB VARCHAR(15),
	@DILIC VARCHAR(20),
	@DISVR VARCHAR(15),
	@DBTYPE INT,
	@DIUSER VARCHAR(10),
	@DIPWD VARCHAR(15),
	@DBID VARCHAR(10),
	@DBPWD VARCHAR(15),
	@DITURST BIT = 0 ,
    @newPassword VARCHAR(100) = 'Ves?147369',  -- New parameter for new password when resetting
    @FunctionName VARCHAR(50) = 'userunlock',     -- userunlock ,setpassword ,COAUnblock，COAblock
	 @AccountCode VARCHAR(50)= NULL,
	 @isacctunblock BIT = 1 ,
	@Result INT OUTPUT ,
	@errorDescription NVARCHAR(255) OUTPUT
	with encryption
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @object INT
    DECLARE @hr INT,@userid INT
    DECLARE @error INT
    DECLARE @errorMsg NVARCHAR(255)
    DECLARE @companyName NVARCHAR(100)
    DECLARE @docEntry INT
    DECLARE @poObject INT
    DECLARE @POCancel INT
    DECLARE @recordset INT 
    DECLARE @retVal INT
    DECLARE @newDocEntry INT
	DECLARE @query NVARCHAR(255)
    DECLARE @errorCode INT
    DECLARE @userManager INT    -- For user management operations
	DECLARE @username VARCHAR(150)
	DECLARE @EoF BIT
    -- 初始化
    SET @hr = 0
    SET @errorDescription = ''
    
    -- 创建 SAPbobsCOM.Company 对象
    EXEC @hr = sp_OACreate 'SAPbobsCOM.Company', @object OUT
    IF @hr <> 0
    BEGIN
        EXEC @hr = sp_OAGetErrorInfo @object, @error OUT, @errorMsg OUT
        SET @errorDescription = N'Step 2: Fail to create Company object: ' + ISNULL(@errorMsg, 'Unknown')
		set @Result = -1
        GOTO End_Routine
    END

    -- 设置连接参数
    EXEC @hr = sp_OASetProperty @object, 'Server', @DISVR
    EXEC @hr = sp_OASetProperty @object, 'SLDServer',@DILIC
    EXEC @hr = sp_OASetProperty @object, 'DbServerType',@DBTYPE 
    EXEC @hr = sp_OASetProperty @object, 'CompanyDB', @CompanyDB
    EXEC @hr = sp_OASetProperty @object, 'UserName', @DIUSER
    EXEC @hr = sp_OASetProperty @object, 'Password', @DIPWD
    EXEC @hr = sp_OASetProperty @object, 'DbUserName', @DBID
    EXEC @hr = sp_OASetProperty @object, 'DbPassword', @DBPWD
    EXEC @hr = sp_OASetProperty @object, 'UseTrusted', @DITURST

    -- 连接到公司数据库
    EXEC @hr = sp_OAMethod @object, 'Connect', @retVal OUT
    IF @retVal <> 0 AND @hr = 0 
    BEGIN
        EXEC @hr = sp_OAMethod @object, 'GetLastErrorDescription', @errorDescription OUT
        SET @errorDescription = 'Step 3: Database connection failed - ' + @errorDescription
		set @Result = -1
       GOTO Cleanup
    END
   
   	   -- 创建 Recordset 对象用于查询
    EXEC @hr = sp_OAMethod @object,'GetBusinessObject', @recordset OUT,300
	EXEC @hr = sp_OAMethod @object, 'GetBusinessObject', @userManager OUT,12
   
    -- 添加用户解锁功能
	
    IF @FunctionName = 'userunlock'
    BEGIN
    -- 使用 DoQuery 查询
    SET @query = 'select userid from ousr where   Locked = ''Y'' AND  user_code=''' +@USER+''''
    EXEC @hr = sp_OAMethod @recordset, 'DoQuery', NULL, @query
	 -- 获取查询结果
	 EXEC @hr = sp_OAGetProperty @recordset, 'EoF', @EoF OUT
	 IF @EoF= 0 
	 BEGIN
		 EXEC @hr = sp_OAGetProperty @recordset, 'Fields.Item(0).Value', @userid OUT
		EXEC @hr = sp_OAMethod @userManager, 'GetByKey', @retVal OUT, @userid
		IF  @retVal <> 0
        BEGIN
            -- 解锁用户
			EXEC @hr = sp_OAGetProperty @userManager, 'UserName', @username OUT
			SET @username = replace(REPLACE(@username,'Locked',''),'-','')
			EXEC @hr = sp_OASetProperty @userManager, 'UserName', @username
            EXEC @hr = sp_OASetProperty @userManager, 'locked', 0
			 EXEC @hr = sp_OAMethod @userManager, 'Update', @retVal OUT
            IF  @retVal <> 0
            BEGIN
			   EXEC @hr = sp_OAMethod @object, 'GetLastErrorDescription', @errorDescription OUT
                SET @errorDescription = 'Failed to unlock user '+@errorDescription
				set @Result = -1
                GOTO Cleanup
            END
			ELSE 
			  BEGIN
			   set @Result =0 
			   set  @errorDescription = 'User Unlocked'
			  END 
        END
     END
	 Else
	 Begin
	   SET @errorDescription = 'SAP B1 User not found or not locked'
	   set @Result = -1
	    GOTO Cleanup
	 End
   end 

    -- 添加密码重置功能
   Else IF @FunctionName = 'setpassword'
    BEGIN
       -- 使用 DoQuery 查询
    SET @query = 'select userid from ousr where  user_code=''' +@USER+''''
    EXEC @hr = sp_OAMethod @recordset, 'DoQuery', NULL, @query
	 -- 获取查询结果
	 EXEC @hr = sp_OAGetProperty @recordset, 'EoF', @EoF OUT
	 IF @EoF= 0 
	 BEGIN
		 EXEC @hr = sp_OAGetProperty @recordset, 'Fields.Item(0).Value', @userid OUT
		EXEC @hr = sp_OAMethod @userManager, 'GetByKey', @retVal OUT, @userid
		IF  @retVal <> 0
        BEGIN
            -- 
            EXEC @hr = sp_OASetProperty @userManager, 'UserPassword',@newPassword
			 EXEC @hr = sp_OAMethod @userManager, 'Update', @retVal OUT
            IF  @retVal <> 0
            BEGIN
			   EXEC @hr = sp_OAMethod @object, 'GetLastErrorDescription', @errorDescription OUT
                SET @errorDescription = 'failed to reset password ：'+@errorDescription
				set @Result = -1
                GOTO Cleanup
            END
			ELSE 
			  BEGIN
			   set  @errorDescription = 'your SAPB1 password is : '+@newPassword
			     set @Result = 0 
			  END 
        END
    END
	else 
    begin
	 SET @errorDescription = 'SAP B1 User not found '
				set @Result = -1
                GOTO Cleanup
	end  
   end 
     -- 添加CoA unblock
   Else IF @FunctionName = 'COAUnblock'
     Begin 
       declare  @CoA INT
	   declare @blkmanpost int 
	   SET @query = 'SELECT 1 FROM OACT t WHERE  t.AcctCode =''' +@AccountCode+''''
    EXEC @hr = sp_OAMethod @recordset, 'DoQuery', NULL, @query
	 -- 获取查询结果
	 EXEC @hr = sp_OAGetProperty @recordset, 'EoF', @EoF OUT
	 IF @EoF <> 0 
	 BEGIN
	    SET @errorDescription = 'code not found '
				set @Result = -1
                GOTO Cleanup
	 end 
	 -- SET @query = 'SELECT 1 FROM OUSR t0 INNER JOIN USR7 t1 ON t0.USERID= t1.UserId  INNER JOIN OUGR t2 ON t1.GroupId = t2.GroupId WHERE t0.Locked = ''N'' AND t0.GROUPS <> 99 AND (t2.GroupName IN (''Finance-JV'',''Finance-JEPostings'') or SUPERUSER=''Y'' ) AND t0.USER_CODE=''' +@USER+''''
  SET @query = 'SELECT  MAX(iif(t0.SUPERUSER = ''Y'',1,0) + IIF(t2.GroupName in (''Finance-JV'',''Finance-JEPostings''),1,0)) FROM OUSR t0 INNER JOIN USR7 t1 ON t0.USERID = t1.UserId INNER JOIN OUGR t2 ON t1.GroupId = t2.GroupId WHERE t0.USER_CODE =''' +@USER+''''
  EXEC @hr = sp_OAMethod @recordset, 'DoQuery', NULL, @query
	 -- 获取查询结果
	 EXEC @hr = sp_OAGetProperty @recordset, 'EoF', @EoF OUT
	 IF @EoF <> 0 
	 BEGIN
	    SET @errorDescription = 'Your sapb1 userid does not have access to Journal Vouchers or Journal Entry.'
				set @Result = -1
                GOTO Cleanup
	 end 
	 
	-- 获取 Chart of Accounts 对象 (类型1)
	EXEC @hr = sp_OAMethod @object, 'GetBusinessObject', @CoA OUT, 1
	IF @hr <> 0 BEGIN
		EXEC sp_OAGetErrorInfo @object, @errorDescription OUT
		SET @errorDescription = 'Failed to get ChartOfAccount Object'
		set @Result = -1
		GOTO Cleanup
	END
    if @isacctunblock = 1 
	begin 
	 SET @query = 'SELECT 1 FROM OACT t WHERE t.BlocManPos=''y'' AND len(t.AcctCode)<>0 and  t.AcctCode =''' +@AccountCode+''''
    EXEC @hr = sp_OAMethod @recordset, 'DoQuery', NULL, @query
	 -- 获取查询结果
	 EXEC @hr = sp_OAGetProperty @recordset, 'EoF', @EoF OUT
	 IF @EoF <> 0 
	 BEGIN
	    SET @errorDescription = 'code was not be blocked for manual post'
				set @Result = -1
                GOTO Cleanup
	 end 
	 end 
	 else 
	  Begin
	      SET @query = 'SELECT 1 FROM OACT t WHERE t.BlocManPos=''n'' AND len(t.AcctCode)<>0 and  t.AcctCode =''' +@AccountCode+''''
    EXEC @hr = sp_OAMethod @recordset, 'DoQuery', NULL, @query
	 -- 获取查询结果
	 EXEC @hr = sp_OAGetProperty @recordset, 'EoF', @EoF OUT
	 IF @EoF <> 0 
	 BEGIN
	    SET @errorDescription = 'code already has been blocked for manual post'
				set @Result = -1
                GOTO Cleanup
	 end 
	   SET @query = 'SELECT 1 FROM OUSR t0 where t0.Locked = ''N'' AND t0.GROUPS <> 99 AND  SUPERUSER=''Y''  AND t0.USER_CODE=''' +@USER+''''
    EXEC @hr = sp_OAMethod @recordset, 'DoQuery', NULL, @query
	 -- 获取查询结果
	 EXEC @hr = sp_OAGetProperty @recordset, 'EoF', @EoF OUT
	 IF @EoF <> 0 
	 BEGIN
	    SET @errorDescription = 'your SAPB1 userid is not SUPERUSER role to update GL account '
				set @Result = -1
                GOTO Cleanup
	 end 
	  end 
	EXEC @hr = sp_OAMethod @CoA, 'GetByKey', @retVal OUT, @AccountCode
	-- 检查账户是否存在
	IF @retVal <> 0  -- 成功找到账户
	BEGIN
	   if @isacctunblock = 1 
	     Begin
		EXEC @hr = sp_OASetProperty @CoA, 'BlockManualPosting', 0
		end 
	   else 
	     Begin
		   EXEC @hr = sp_OASetProperty @CoA, 'BlockManualPosting', 1
		 end 
		-- 执行更新操作
		EXEC @hr = sp_OAMethod @CoA, 'Update', @retVal OUT
		IF  @retVal <> 0 BEGIN
			EXEC @hr = sp_OAMethod @object, 'GetLastErrorDescription', @errorDescription OUT
			SET @errorDescription = 'Faid to update : ' + @errorDescription
			set @Result = -1
			GOTO Cleanup
		END
		
		-- 更新成功
		if @isacctunblock = 1 
		SET @errorDescription = 'code unblocked'
		Else 
		 Begin
		 SET @errorDescription = 'code blocked '
		 end 
		SET @Result = 0
	END
	ELSE  -- 未找到账户
	BEGIN
		SET @errorDescription = 'code Not Found'
		SET @Result = -1
		GOTO Cleanup
	end 
  END
 ELSE
        RAISERROR('InValid FunctionName', 16, 1);
    -- 清理资源
	Cleanup:
	BEGIN 
    IF @userManager IS NOT NULL
    BEGIN
        EXEC sp_OADestroy @userManager
    END
       IF @recordset IS NOT NULL
    BEGIN
        EXEC sp_OADestroy @recordset
    END
	IF @CoA IS NOT NULL EXEC sp_OADestroy @CoA
	if @object is NOT NULL
	BEGIN
      EXEC @hr = sp_OAMethod @object, 'Disconnect', NULL
      EXEC sp_OADestroy @object
	end 
	END 
  End_Routine:
	  RETURN;
END