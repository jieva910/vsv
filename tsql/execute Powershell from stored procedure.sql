USE [SAPB1_CS_TST]
GO
/****** Object:  StoredProcedure [dbo].[SBO_SP_PostTransactionNotice]    Script Date: 9/18/2022 12:08:15 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[SBO_SP_PostTransactionNotice]

@object_type varchar(20), 				-- SBO Object Type
@transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
@num_of_cols_in_key int,
@list_of_key_cols_tab_del varchar(255),
@list_of_cols_val_tab_del varchar(255)

AS

begin

-- Return values
declare @error  int				-- Result (0 for no error)
declare @error_message varchar (200) 		-- Error string to be displayed
select @error = 0
select @error_message = N'Ok'

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE

--------------------------------------------------------------------------------------------------------------------------------
IF @object_type = '13' AND @transaction_type IN ('A','U')
BEGIN
  declare @Ps varchar(8000),@dbsvr varchar(50),@db varchar(50),@sldsvr varchar(50),@ScriptString varchar(4000)
	-- get SLD Serever of SAPB1
	select @sldsvr=lsrv from [SBO-COMMON].dbo.SLIC

	--select SYSTEM_USER
	--select @@SERVERNAME
	--select DB_NAME()

	set @Ps =  'powershell.exe -WindowStyle Hidden -NonInteractive -NoProfile -Command $cmp = New-Object -ComObject ''SAPBOBSCom.company'';'+
		 ' function SendeMail($content,$sub,$to) { Send-MailMessage -Body $content -From SAPB1APP@Vsv.COM  -To $to -SmtpServer ''APMailrelay.vesuvius.com'' -port 25 -Subject $sub };'+
		 ' $cmp.SLDServer = '''+@sldsvr+''';	$cmp.Server = '''+@@SERVERNAME+''';	$cmp.CompanyDB = '''+DB_NAME()+''';$cmp.DbServerType = 8;	$cmp.UseTrusted = $true;'+
		 '	$cmp.UserName = ''corp\jieva'';$cmp.Password =''Vesint-999'';$cmp.Connect();'+
		 ' 	if (!$cmp.Connected){SendeMail $cmp.GetLastErrorDescription()  ''Company Connection Failure'' ''evan.ji@vesuvius.com''; exit }; $obj = $cmp.GetBusinessObject(13);'+
		 ' if ($obj.GetByKey('+@list_of_cols_val_tab_del+')){$obj.Printed = 1 ;$obj.Update()}; $cmp.Disconnect() '
	EXEC master..xp_cmdshell @Ps,'no_output'
 --- had performance issue 
END
-- Select the return values
select @error, @error_message

end



---- version 2 call sql job in writing with powershell 
USE [SAPB1_CS_TST]
GO
/****** Object:  StoredProcedure [dbo].[SBO_SP_PostTransactionNotice]    Script Date: 9/18/2022 2:06:33 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[SBO_SP_PostTransactionNotice]

@object_type varchar(20), 				-- SBO Object Type
@transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
@num_of_cols_in_key int,
@list_of_key_cols_tab_del varchar(255),
@list_of_cols_val_tab_del varchar(255)

AS

begin

-- Return values
declare @error  int				-- Result (0 for no error)
declare @error_message varchar (200) 		-- Error string to be displayed
select @error = 0
select @error_message = N'Ok'

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE

--------------------------------------------------------------------------------------------------------------------------------
IF @object_type = '13' AND @transaction_type IN ('A','U')
BEGIN
  declare @Ps varchar(8000),@dbsvr varchar(50),@db varchar(50),@sldsvr varchar(50),@ScriptString varchar(4000)
  declare @ps2 varchar(8000)
	-- get SLD Serever of SAPB1
	select @sldsvr=lsrv from [SBO-COMMON].dbo.SLIC

	--select SYSTEM_USER
	--select @@SERVERNAME
	--select DB_NAME()
	set @Ps =  ' $cmp = New-Object -ComObject ''SAPBOBSCom.company'';'+
		 ' function SendeMail($content,$sub,$to) { Send-MailMessage -Body $content -From SAPB1APP@Vsv.COM  -To $to -SmtpServer ''APMailrelay.vesuvius.com'' -port 25 -Subject $sub };'+
		 ' $cmp.SLDServer = '''+@sldsvr+''';	$cmp.Server = '''+@@SERVERNAME+''';	$cmp.CompanyDB = '''+DB_NAME()+''';$cmp.DbServerType = 8;	$cmp.UseTrusted = $true;'+
		 '	$cmp.UserName = ''corp\jieva'';$cmp.Password =''Vesint-999'';[void]$cmp.Connect();'+
		 ' 	if (!$cmp.Connected){SendeMail $cmp.GetLastErrorDescription()  ''Company Connection Failure'' ''evan.ji@vesuvius.com''; exit }; $obj = $cmp.GetBusinessObject(13);'+
		 ' if ($obj.GetByKey('+@list_of_cols_val_tab_del+')){$obj.Printed = 1 ;[void]$obj.Update()};$cmp.Disconnect(); exit '
	    EXEC msdb.dbo.sp_update_jobstep    @job_name=N'tst', @step_name=N'print', 
            @step_id=1, 
            @cmdexec_success_code=0, 
            @on_success_action=1, 
            @on_fail_action=2, 
            @retry_attempts=0, 
            @retry_interval=0, 
            @os_run_priority=0, @subsystem=N'POWERSHELL', 
            @command=@Ps, 
            @flags=0

    waitfor delay '00:00:00:03'

	exec msdb.dbo.sp_start_job @job_name ='tst'
	
END
-- Select the return values
select @error, @error_message

end