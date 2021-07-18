/****** Object:  LinkedServer [VTECRMPROD]    Script Date: 02/24/2017 16:07:44 ******/
EXEC master.dbo.sp_addlinkedserver @server = N'VTECRMPROD', @srvproduct=N'MySQL', @provider=N'MSDASQL', @datasrc=N'VTECRM PROD', @provstr=N'DRIVER={MySQL ODBC 5.3 ANSI Driver};SERVER=172.19.1.2;PORT=3306;DATABASE=vte;USER=italcom;PASSWORD=italcom;OPTION=3;', @catalog=N'vte'
 /* For security reasons the linked server remote logins password is changed with ######## */
EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'VTECRMPROD',@useself=N'False',@locallogin=NULL,@rmtuser=N'italcom',@rmtpassword='########'

GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'collation compatible', @optvalue=N'false'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'data access', @optvalue=N'true'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'dist', @optvalue=N'false'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'pub', @optvalue=N'false'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'rpc', @optvalue=N'true'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'rpc out', @optvalue=N'true'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'sub', @optvalue=N'false'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'connect timeout', @optvalue=N'0'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'collation name', @optvalue=null
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'lazy schema validation', @optvalue=N'false'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'query timeout', @optvalue=N'0'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'use remote collation', @optvalue=N'true'
GO

EXEC master.dbo.sp_serveroption @server=N'VTECRMPROD', @optname=N'remote proc transaction promotion', @optvalue=N'true'
GO


