USE [KnoS_ZSI]
GO

/****** Object:  View [dbo].[Metodo_View_PDFDocumenti]    Script Date: 10/27/2017 12:34:20 ******/
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[Metodo_View_PDFDocumenti]'))
DROP VIEW [dbo].[Metodo_View_PDFDocumenti]
GO

USE [KnoS_ZSI]
GO

/****** Object:  View [dbo].[Metodo_View_PDFDocumenti]    Script Date: 10/27/2017 12:34:20 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[Metodo_View_PDFDocumenti]
AS

select 
od.IdObject 
, od.IdDoc
, od.FileName collate Latin1_General_CI_AS as FileName
, od.DataFileName collate Latin1_General_CI_AS as DataFileName
, od.DateVersion 
from object_doc od 
GO


grant all on [Metodo_View_PDFDocumenti] to metodo98
go


DECLARE @userid as Nvarchar(500)
DECLARE @sql as Nvarchar(500)
DECLARE db_cursor CURSOR FOR  
SELECT userid 
FROM zsi.dbo.tabutenti where nrterminale > 0 
--AND USERID = 'TSMAGAZZINO1'

OPEN db_cursor   
FETCH NEXT FROM db_cursor INTO @userid   

WHILE @@FETCH_STATUS = 0   
BEGIN   
		SET @sql = 'EXEC sp_revokedbaccess N''' + @userid + ''''
		EXECUTE SP_EXECUTESQL @SQL

		SET @sql = 'EXEC sp_grantdbaccess N''' + @userid + ''', N''' + @userid + ''''
		EXECUTE SP_EXECUTESQL  @SQL

	  SET @sql = 'EXEC sp_addrolemember N''Metodo98'', N''' + @userid + ''''
	--EXEC sp_grantdbaccess N'TsMagazzino1', N'TsMagazzino1'
	--EXEC sp_addrolemember N'TsMagazzino1'
		EXECUTE SP_EXECUTESQL  @SQL

       FETCH NEXT FROM db_cursor INTO @userid  
END   

CLOSE db_cursor   
DEALLOCATE db_cursor
