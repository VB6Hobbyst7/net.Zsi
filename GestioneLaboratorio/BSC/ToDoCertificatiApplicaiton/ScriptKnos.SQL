-- scrpt KNOS
/****** Object:  View [dbo].[Metodo_View_LinkageBollettini]    Script Date: 21/05/2015 18:22:26 ******/
DROP VIEW [dbo].[Metodo_View_LinkageBollettini]
GO

/****** Object:  View [dbo].[Metodo_View_LinkageBollettini]    Script Date: 21/05/2015 18:22:26 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[Metodo_View_LinkageBollettini]
AS

select 
	ol.idparent AS IDOBJECT_BOL 
	, ol.IdChild AS IDOBJECT_ART
	, olcli.IdChild AS IDOBJECT_CLIFOR
	, od.*
from object_linkage ol 
inner join Object_Linkage olcli on ol.IdParent = olcli.IdParent and ol.idchild <> olcli.IdChild
inner join object_doc od on od.idobject = ol.IdParent
where ol.idattr = 5036 and olcli.IdAttr = 19
and ol.idparent in (select idobject from object_object where idclass = 5013)
--AND OD.Version = OD.CurrentVersion


GO


/****** Object:  View [dbo].[[Metodo_View_LinkageSchedeTecniche]]    Script Date: 21/05/2015 18:22:26 ******/
DROP VIEW [dbo].[Metodo_View_LinkageSchedeTecniche]
GO

/****** Object:  View [dbo].[[Metodo_View_LinkageSchedeTecniche]]    Script Date: 21/05/2015 18:22:26 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[Metodo_View_LinkageSchedeTecniche]
AS

select 
	ol.idparent AS IDOBJECT_SCH 
	, ol.IdChild AS IDOBJECT_ART
	, olcli.IdChild AS IDOBJECT_CLIFOR
	, od.*
from object_linkage ol 
inner join Object_Linkage olcli on ol.IdParent = olcli.IdParent and ol.idchild <> olcli.IdChild
inner join object_doc od on od.idobject = ol.IdParent
where ol.idattr = 5036 and olcli.IdAttr = 19
and ol.idparent in (select idobject from object_object where idclass = 5012)
--AND OD.Version = OD.CurrentVersion


GO



