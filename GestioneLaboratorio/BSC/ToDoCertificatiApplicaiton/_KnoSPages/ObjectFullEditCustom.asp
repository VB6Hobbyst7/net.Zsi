
<!-- #include virtual="knos/system/include/incConst.asp" -->
<%
	'Redirezione su accesso negato
	AspRedirectLogin = URL_LOGIN_SITE &"&TargetUrl=DlgObjectFullEdit.asp?"& Request.QueryString
	AspRedirectDenied = URL_DENIED_ACCESS
	AspLoadUserProfile = True
%>
<!-- #include virtual="KnoS/system/include/incValidateCookie.asp" -->
<!-- #include virtual="knos/system/ICKnos2.asp" -->
<!-- #include virtual="knos/system/env/lib/asp/AspQuery.asp" -->
<!-- #include virtual="knos/system/env/lib/asp/KnoS.asp" -->
<!-- #include virtual="/custom/incCustomSearch.asp" -->

<%
'Raccolta dei permessi dell'utente corrente
Initialize

Set o = objApplication.Objects.Item(AspIdObject)
If IsEmpty(o) Or Err.number <> 0 Then
	Response.Redirect AspRedirectDenied
End If

If Not CheckCustomObjectIsVisible(objApplication, o) Then
	Response.Redirect AspRedirectDenied
End If

AspIdClass = o.IdClass
AspIdStatus = o.IdStatus
AspSuffix = "\"& AspIdClass &"\"& AspIdStatus

AspPermission_Read = objApplication.Security.Operations.Check( CStr(OP_KNOS_CLASS_READ & AspSuffix), CLng(0) )
AspPermission_Edit = objApplication.Security.Operations.Check( CStr(OP_KNOS_CLASS_EDIT & AspSuffix), CLng(0) )
AspPermission_Doc = objApplication.Security.Operations.Check( CStr(OP_KNOS_CLASS_UPLOAD & AspSuffix), CLng(0) )
AspPermission_Link = objApplication.Security.Operations.Check( CStr(OP_KNOS_CLASS_LINK & AspSuffix), CLng(0) )
AspPermission_Permission = objApplication.Security.Operations.Check( CStr(OP_KNOS_CLASS_PERMISSION & AspSuffix), CLng(0) )
AspPermission_Action = objApplication.Security.Operations.Check( CStr(OP_KNOS_CLASS_ACTION & "\"& AspIdClass &"\%"), CLng(0) )
AspPermission_Catalog = CheckCustomObjectCatalogEnabled(objApplication, AspIdObject, AspIdClass, AspIdStatus)
AspPermission_Catalog_Template = objApplication.Security.Operations.Check( CStr(OP_KNOS_CATALOG_TEMPLATE), CLng(0) )

'Tab azioni disabilitato in caso di MatchTemplate
AspAction = Request.QueryString("Action")
If LCase(CStr(AspAction)) = "matchtemplate" Then
	 AspPermission_Action = false
End If

'Su richiesta si visualizzano solo i tab elencati in QueryString
AspTab = AspQuery.requestQueryString("Tab", "")
If AspTab <> "" Then
	AspPermission_Edit = AspPermission_Edit And InStr(1, AspTab, "edit", 1) > 0
	AspPermission_Doc = AspPermission_Doc And InStr(1, AspTab, "upload", 1) > 0
	AspPermission_Link = AspPermission_Link And InStr(1, AspTab, "link", 1) > 0
	AspPermission_Permission = AspPermission_Permission And InStr(1, AspTab, "permission", 1) > 0
	AspPermission_Action = AspPermission_Action And InStr(1, AspTab, "action", 1) > 0
	AspPermission_Catalog = AspPermission_Catalog And InStr(1, AspTab, "catalog", 1) > 0
End If

If (Not AspPermission_Edit) And (Not AspPermission_Doc) And (Not AspPermission_Link) And (Not AspPermission_Permission) And (Not AspPermission_Action) And (Not AspPermission_Catalog) Then
	Response.Redirect "/knos/web/PrintObject.asp?"& Request.QueryString &"&check=all"
Else
	AspPermission_Property = true
End If

'Check out ritardato se possibile
If AspIdClass=0 Or AspIdStatus=0 Then
	Err.Clear
	o.CheckOut
	If Err.number <> 0 Then
		Response.Redirect "/knos/web/PrintObject.asp?"& Request.QueryString &"&check=all&lock=1"
	End If
End If

Dim AspIsAdministrator : AspIsAdministrator = LCase( IsAdministrator() )
Finalize

%>

<HTML>
<HEAD>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8">

<TITLE> <%= Language.label("Modifica pubblicazione") %>&nbsp;<%= AspIdObject %> </TITLE>

<%= AspLinkStylesheet() %>

<STYLE>

.Dlg-clsTABLE
{
	width: 100%;
	height: 100%;
}

.Dlg-clsTD-bottom
{
	height: 20;
}

.Dlg-clsDIV-top
{
	width: 100%;
	height: 200;
}

.Dlg-clsDIV-table
{
	width: 100%;
	height: 100%;
	overflow: auto;
}

.CFormFilter-clsTD-attributes { text-align: left; padding: 2; font-weight: bold; border-bottom-width: 1; vertical-align: top; }
.CTabBar-clsTABLE { border-width: 0; }

</STYLE>

</HEAD>


<BODY>

<%= AspScriptLibrary() %>
<!-- #include virtual="/custom/incCustomTop.asp" -->
<!-- #include virtual="/custom/incCustom.asp" -->
<!-- #include virtual="/knos/web/incFileEdit.asp" -->

<script language="JavaScript">
<!--

var Soap = new CSoap();

var AspIdObject = "<%= AspIdObject %>";
var AspIdClass = "<%= AspIdClass %>";
var AspIdStatus = "<%= AspIdStatus %>";
var AspIdUser = "<%= AspIdUser %>";
var AspIsAdministrator = <% If AspIsAdministrator Then %> true <% Else %> false <% End If %>;
var URL_WEB_SERVICES = "<%= URL_WEB_SERVICES %>";
var DEFAULT_USER_VIEW = "<%= DEFAULT_USER_VIEW %>";
var AspSiteQuery = "<%= AspSiteQuery %>";
var AspOpener = "<%= Request.QueryString("Opener") %>".toLowerCase();
var AspAction = "<%= AspAction %>".toLowerCase();
//	I valori gestiti di AspAction sono:
//		<empty>			gestione ordinaria
//		MatchTemplate	Avanti/Rinuncia con assegnazione catalogo e stato del template
//		Attach			Outlook/Office quando allegano file ad una pubblicazione esistente 
//
var AspPermission_Property = <% If AspPermission_Property Then %> true <% Else %> false <% End If %>;
var AspPermission_Edit = <% If AspPermission_Edit Then %> true <% Else %> false <% End If %>;
var AspPermission_Permission = <% If AspPermission_Permission Then %> true <% Else %> false <% End If %>;

var FormEdit = new CFormFilter("FormEdit", "<%= URL_WEB_SERVICES %>/Object_Form.asp<%= AspSiteQuery %>");
FormEdit.enableDefaultCallback();
var Status =
{
	Minimized : false,
	MinimizedHeight : "200px",
	Height : window.dialogHeight,
	Top : window.dialogTop,
	Left : window.dialogLeft
}

var Splitter = new CSplitter("Splitter");

function Init()
{
	Splitter.add("Doc_Splitter_idDIV", false, "DocOuter_idDIV", "Doc_idDIV", "Doc_Permission_idDIV");

	// Pulsanti:
	var btnsSave = document.all["Save_idBUTTON"].style;
	var btnsSaveExit = document.all["SaveExit_idBUTTON"].style;
	var btnsCancel = document.all["Cancel_idBUTTON"].style;
	var btnsNotify = document.all["Notify_idBUTTON"].style;
	var btnsClose = document.all["Close_idBUTTON"].style;
	var btnsNext = document.all["Next_idBUTTON"].style;
	var btnsUndo = document.all["Undo_idBUTTON"].style;

	switch (AspAction)
	{
		case "matchtemplate":
			btnsSave.display = "none";
			btnsSaveExit.display = "none";
			btnsNotify.display = "none";
			btnsCancel.display = "none";
			btnsClose.display = "none";
			btnsNext.display = "inline";
			btnsUndo.display = "inline";
			break;

		default:
			btnsSave.display = "inline";
			btnsSaveExit.display = "inline";
			btnsNotify.display = "inline";
			btnsCancel.display = "inline";
			btnsClose.display = "inline";
			btnsNext.display = "none";
			btnsUndo.display = "none";	
	}

	if (AspPermission_Property)
	{
		TabBar.enableTab(TAB_PROPERTY, true, false);
		PropertyRefresh();
	}
	if (AspPermission_Edit)
		TabBar.enableTab(TAB_ATTRIBUTE, true, false);

	<% If AspPermission_Doc Then %>
		TabBar.enableTab(TAB_DOC, true, false);
	<% End If %>
	<% If AspPermission_Link Then %>
		TabBar.enableTab(TAB_LINK, true, false);
	<% End If %>
	<% If AspPermission_Permission Then %>
		TabBar.enableTab(TAB_PERMISSION, true, false);
		LoadSubjectSelection();
	<% End If %>
	<% If AspPermission_Action Then %>
		TabBar.enableTab(TAB_ACTION, true, false);
		TabBar.enableTab(TAB_WORKFLOW, true, false);
	<% End If %>
	<% If AspPermission_Catalog Then %>
		TabBar.enableTab(TAB_CATALOG, true, false);
	<% End If %>
	if (AspPermission_Edit)
		TabBar.selected = 1;
	TabBar.refresh();
	FormEdit.soap.setField("IdObject", AspIdObject);
	FormEdit.queryForm();

	// Eventuali valorizzazioni custom (ad es. dei cataloghi)
	CustomObjectFullEdit(AspOpener, FormEdit);
}

//-->
</script>


	<TABLE CLASS="Dlg-clsTABLE" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD CLASS="ColorOutset" STYLE="padding: 0; padding-top: 1; border-width: 0;">

<script language="JavaScript">
<!--

var TAB_PROPERTY = Language.label("Proprietà");
var TAB_ATTRIBUTE = Language.label("Attributi");
var TAB_DOC = Language.label("Documenti");
var TAB_LINK = Language.label("Riferimenti");
var TAB_PERMISSION = Language.label("Destinatari");
var TAB_ACTION = Language.label("Azioni");
var TAB_WORKFLOW = Language.label("Workflow");
var TAB_CATALOG = Language.label("Cataloghi");
EditTab = TAB_ATTRIBUTE;
EditAdd = false;
//var TabBar = new CTabBar("TabBar", [TAB_PROPERTY, TAB_ATTRIBUTE, TAB_DOC, TAB_LINK, TAB_PERMISSION, TAB_ACTION, TAB_WORKFLOW, TAB_CATALOG]);
var TabBar = new CTabBar("TabBar", [TAB_PROPERTY, TAB_ATTRIBUTE, TAB_DOC, TAB_WORKFLOW]);
for (var i=0; i<TabBar.tab.length; i++)
{
	switch (TabBar.tab[i].caption)
	{
		default:
			TabBar.tab[i].enabled = false;
	}
}
TabBar.selected = -1;

TabBar.writeHTML_START(1);

//-->
</script>


<script language="JavaScript">
<!--
	TabBar.writeHTML_START(2);
//-->
</script>

<!---------------------------- TAB PROPERTY START --------------------------->

<TABLE ID="TabProperty_idTABLE" CLASS="CTabBar-clsTABLE-inner" CELLSPACING=0 CELLPADDING=0 STYLE="display: none; border-width: 0;">
<TR><TD>
	<DIV STYLE="overflow: auto; width: 100%; height: 100%;">
		<TABLE CLASS="CPreview-clsTABLE" style="table-layout: auto;" ID="Property_idTABLE">
		<col width="1%" />
		<TR>
			<TD CLASS="CPreview-clsTD-label" NOWRAP><%= Language.label("col: Id") %></TD>
			<TD CLASS="CPreview-clsTD-value" id="Property_IdObject_TD">&nbsp;</TD>
		</TR>
		<TR>
			<TD CLASS="CPreview-clsTD-label" NOWRAP><%= Language.label("col: Tipologia") %></TD>
			<TD CLASS="CPreview-clsTD-value" id="Property_ClassName_TD">&nbsp;</TD>
		</TR>
		<TR>
			<TD CLASS="CPreview-clsTD-label" NOWRAP><%= Language.label("col: Stato") %></TD>
			<TD CLASS="CPreview-clsTD-value" id="Property_StatusName_TD">&nbsp;</TD>
		</TR>
		<TR>
			<TD CLASS="CPreview-clsTD-label" NOWRAP><%= Language.label("Data creazione") %></TD>
			<TD CLASS="CPreview-clsTD-value" id="Property_DateCreation_TD">&nbsp;</TD>
		</TR>
		<TR>
			<TD CLASS="CPreview-clsTD-label" NOWRAP><%= Language.label("Creato da") %></TD>
			<TD CLASS="CPreview-clsTD-value" id="Property_OwnerSubject_TD">&nbsp;</TD>
		</TR>
		<TR>
			<TD CLASS="CPreview-clsTD-label" NOWRAP><%= Language.label("Data modifica") %></TD>
			<TD CLASS="CPreview-clsTD-value" id="Property_DateModify_TD">&nbsp;</TD>
		</TR>
		<TR>
			<TD CLASS="CPreview-clsTD-label" NOWRAP><%= Language.label("Lingua") %></TD>
			<TD CLASS="CPreview-clsTD-value" id="Property_Language_TD">&nbsp;</TD>
		</TR>
		</TABLE>
	</DIV>
</TD></TR>
</TABLE>

<!---------------------------- TAB PROPERTY END --------------------------->

<!---------------------------- TAB ATTRIBUTI START --------------------------->

<TABLE ID="TabAttribute_idTABLE" CLASS="CTabBar-clsTABLE-inner" CELLSPACING=0 CELLPADDING=0 STYLE="display: block; border-width: 0;">
<TR><TD>
	<DIV STYLE="overflow: auto; width: 100%; height: 100%;">

<script language="JavaScript">
<!--
FormEdit.writeHTML();
//-->
</script>

	</DIV>
</TD></TR>
</TABLE>

<!---------------------------- TAB ATTRIBUTI END ----------------------------->

<!--------------------------- TAB DOCUMENTI START ---------------------------->

<TABLE ID="TabDoc_idTABLE" CLASS="CTabBar-clsTABLE-inner" CELLSPACING=0 CELLPADDING=0 STYLE="display: none;">
<TR><TD>
	<DIV id="DocOuter_idDIV" STYLE="width: 100%; height: 100%;">

<script language="JavaScript">
<!--
// TABELLA DOCUMENTI

var DocRecord = new CRecord(["IdDoc", "Url", "FileType", "FileDescr", "FileName", "DataFileName", "DataUrl", "FileProp1", "FileProp2", "FileProp3", "IdUserLock", "UserLock", "LocalPath", "IdVersion", "Version", "Release", "CurrentVersion", "VersionCounter", "DateVersion"]);
var DocTable = new CTable("DocTable", "<%= URL_WEB_SERVICES %>/Documents_Items.asp<%= AspSiteQuery %>");
DocTable.enableCheck = true;
DocTable.view = ViewDoc();
DocTable.urlWeb = "<%= URL_CATALOG %>";
DocTable.pageSize = 0;
DocTable.soap.autoclear = false;
DocTable.moveRecord = { IdObject : AspIdObject, IdDoc : null };
DocTable.moveParam =
{
	 Table : "Object_Doc"
	,Expression : "IdObject={IdObject} AND IdDoc={IdDoc}"
	,Key : ["IdDoc"]
}
DocTable.onquery = function()
{
	document.getElementById("DocPermission_idDIV").style.visibility = "hidden";
	DocPermissionTable.data = {};
}
DocTable.onclick = function(nRow)
{
	this.select(nRow);
	var row = DocTable.row[nRow];
	DocPermissionTable.data =
	{
		 IdObject: row.IdObject
		,IdVersion: row.IdVersion
		,DocTable: this
	}
	DocPermissionTable.soap.setField("IdObject", DocPermissionTable.data.IdObject);
	DocPermissionTable.soap.setField("IdVersion", DocPermissionTable.data.IdVersion);
	DocPermissionTable.query();
	document.getElementById("DocPermission_idDIV").style.visibility = "visible";
}
DocTable.ondblclick = function(nRow)
{
	DocTable.onclick(nRow);
	DoDocEdit(false);
}
//-->
</script>

	<DIV CLASS="Dlg-clsDIV-table" ID="Doc_idDIV" STYLE="background: window; width: 100%; height: 70%; overflow: auto;">
	<script language="JavaScript">
	<!--
	DocTable.writeHTML();
	//-->
	</script>
	</DIV>

	<!-- Doc Splitter -->
	<DIV CLASS="VSplitter-clsDIV-normal" ID="Doc_Splitter_idDIV"></DIV>

	<DIV CLASS="Dlg-clsDIV-table" ID="Doc_Permission_idDIV" STYLE="background: window; width: 100%; height: 30%; overflow: auto;">
		<table cellspacing="0" cellpadding="0" style="width: 100%; height: 100%;">
		<tr>
			<td style="height: 34px; padding-top: 10px;">
			<p class="KnoS-clsFont ColorOutset" style="padding: 2px; padding-left: 4px; font-weight: bold;"><%= Language.label("Documento accessibile solo a:") %></p>
			</td>
		</tr>
		<tr>
			<td style="vertical-align: top;">
			<div id="DocPermission_idDIV" STYLE="background: window; width: 100%; height: 100%; overflow: auto; visibility: hidden;">
			<script language="JavaScript">
			<!--
				var DocPermissionTable = new CTable("DocPermissionTable", "<%= URL_WEB_SERVICES %>/Document_Permission_Items.asp<%= AspSiteQuery %>");
				DocPermissionTable.enableCheck = true;
				DocPermissionTable.view = ViewSubject();
				DocPermissionTable.description = ""
				DocPermissionTable.writeHTML();
			//-->
			</script>
			</div>
			</td>
		</tr>
		</table>
	</DIV>
</DIV>
</TD></TR>
</TABLE>

<!---------------------------- TAB DOCUMENTI END ----------------------------->

<!------------------------------ TAB LINK START ------------------------------>

<TABLE ID="TabLink_idTABLE" CLASS="CTabBar-clsTABLE-inner" CELLSPACING=0 CELLPADDING=0 STYLE="display: none;">
<TR><TD>
	<DIV STYLE="width: 100%; height: 100%;">

<script language="JavaScript">
<!--
// TABELLA LINK
var LinkRecord = new CRecord(["IdLink", "LinkType", "IdObjectTo", "Url", "LinkDescr"]);
var LinkTable = new CTable("LinkTable", "<%= URL_WEB_SERVICES %>/LinksTo_Items.asp<%= AspSiteQuery %>");
LinkTable.enableCheck = true;
LinkTable.autoselect = true;
LinkTable.view = ViewLink();
LinkTable.pageSize = 0;
LinkTable.soap.autoclear = false;
LinkTable.moveRecord = { IdLink : null };
LinkTable.moveParam =
{
	 Table : "Object_Link"
	,Expression : "IdLink={IdLink}"
	,Key : ["IdLink"]
};
LinkTable.ondblclick = function()
{
	DoLinkEdit(false);
}


function DoRadioLinkInterno()
{
	if ( !document.all["RadioLink_Interno"].checked )
		return;
	document.all["FormLink_Label"].innerText = Language.label("Id pubbl.");
	document.all["FormLink_Link"].focus();
	document.all["FormLink_Link"].value = "";
	document.all["Protocollo_idTR"].style.display = "none";
	document.all["Protocol"].style.display = "none";
	document.getElementById("FormLink_Ellipses").style.visibility = "visible";
}

function DoRadioLinkEsterno()
{
	if ( !document.all["RadioLink_Esterno"].checked )
		return;
	document.all["FormLink_Label"].innerText = "Url:";
	document.all["FormLink_Link"].focus();
	document.all["FormLink_Link"].value = "http://"
	document.all["Protocollo_idTR"].style.display = "block";
	document.all["Protocol"].style.display = "block";
	if (document.all["Protocol"].value == "file://")
		document.getElementById("FormLink_Ellipses").style.visibility = "visible";
	else
		document.getElementById("FormLink_Ellipses").style.visibility = "hidden";
}

function DoLinkEllipses(inputTarget)
{
	var target;
	if ( document.all["RadioLink_Interno"].checked )
		WindowOpen("Default.asp<%= AspSiteQuery %>", "LinkedForm_"+ AspIdObject);
	else
		DoBrowse(inputTarget);
}

function DoLinkPaste()
{
	var clip = window.clipboardData.getData("Text");
	var idClip = "KnoS:IdObject:";
	var idObjectToLink;
	var value;
	if ( clip.slice(0, idClip.length) == idClip )
	{
		value = clip.slice(idClip.length);
		if ( document.all["RadioLink_Esterno"].checked )
		{
			document.all["RadioLink_Interno"].checked = true;
			DoRadioInterno();
		}
		if ( !isNaN(parseInt(value, 10)) )
			document.all["FormLink_Link"].value = Trim(value);
	}
	else
	{
		if ( document.all["RadioLink_Esterno"].checked )
			document.all["FormLink_Link"].value = clip;
		else if ( !isNaN(parseInt(clip, 10)) )
			document.all["FormLink_Link"].value = Trim(clip);
	}
	window.event.returnValue = false;
}

function DoLinkSave()
{
	var target = document.all["FormLink_Link"];
	if ( IsEmpty(target.value) )
		return;

	LinkTable.soap.setField("IdObject", AspIdObject);
	if ( document.all["RadioLink_Interno"].checked )
	{
		if ( target.value == AspIdObject )
		{
			window.alert(Language.label("Il riferimento di una pubblicazione a se stessa non è ammesso"));
			return target.focus();
		}
		LinkTable.soap.setField("LinkType", 0);
		LinkTable.soap.setField("LinkIdObject", Trim(target.value));
	}
	else
	{
		LinkTable.soap.setField("LinkType", 1);
		LinkTable.soap.setField("LinkUrl", target.value);
	}
	LinkTable.soap.setField("LinkDescr", Trim(document.all["FormLink_Descr"].value));
	if ( LinkTable.soap.parseResponse("<%= URL_WEB_SERVICES %>/LinksTo_Add.asp<%= AspSiteQuery %>") )
		DoLinkRefresh();
}

function DoLinkCancel()
{
	document.all["FormLink_Link"].value = "";
	document.all["FormLink_Descr"].value = "";
}

function DoLinkRefresh()
{
	LinkTable.soap.setField("IdObject", AspIdObject);
	LinkTable.query();
}

function DoLinkEdit(bChecked)
{
	var bRefresh = false;
	var checkedRecord;
	if (bChecked)
		checkedRecord = LinkTable.checkedRecord(LinkRecord);
	else
		checkedRecord = LinkTable.selectedRecord(LinkRecord);
	var inputText = new Object();
	inputText.title = Language.label("Modifica dati riferimento");
	for (var i=0; i<checkedRecord.length; i++)
	{
		inputText.caption = Language.label("Modifica dati per") + " <B>" + checkedRecord[i].getValue("Url") + "</B>:";
		inputText.IdObject = AspIdObject;
		inputText.LinkType = checkedRecord[i].getValue("LinkType");
		inputText.LinkIdObject = checkedRecord[i].getValue("IdObjectTo");
		inputText.Url = checkedRecord[i].getValue("Url");
		inputText.LinkDescr = checkedRecord[i].getValue("LinkDescr");
		var features = "help:0; resizable:1; status:0; dialogHeight:160px; dialogWidth:500px;";
		if ( window.showModalDialog("/knos/web/DlgObjectLinkEdit.asp<%= AspSiteQuery %>&title=" + inputText.title, inputText, features) == true )
		{
			// Campi immutabili
			LinkTable.soap.setField("IdObject", AspIdObject);
			LinkTable.soap.setField("IdLink", checkedRecord[i].getValue("IdLink"));
			LinkTable.soap.setField("LinkType", checkedRecord[i].getValue("LinkType"));

			// Campi editabili e controllati nel dialogo di edit
			LinkTable.soap.setField("LinkIdObject", inputText.LinkIdObject);
			LinkTable.soap.setField("Url", inputText.Url);
			LinkTable.soap.setField("LinkDescr", inputText.LinkDescr);
			LinkTable.soap.parseResponse("<%= URL_WEB_SERVICES %>/Link_Save.asp<%= AspSiteQuery %>");

			bRefresh = true;
		}
	}
	if ( bRefresh )
	{
		LinkTable.soap.setField("IdObject", AspIdObject);
		LinkTable.query();
	}
}

function DoLinkElimina(bChecked)
{
	var list;
	if (bChecked)
		list = LinkTable.checkedColumn("IdLink");
	else
		list = LinkTable.selectedColumn("IdLink");
	switch (list.length)
	{
		case 0:
			return window.alert(Language.label("Selezionare il riferimento da cancellare"));
		case 1:
			if ( !window.confirm(Language.label("Cancellare il riferimento {0}", (bChecked? Language.label("spuntato") : Language.label("selezionato")))))
				return;
			break;
		default:
			if ( !window.confirm(Language.label("Cancellare i riferimenti {0}", (bChecked? Language.label("spuntati") : Language.label("selezionati")))))
				return;
	}
	LinkTable.soap.setField("IdObject", AspIdObject);
	LinkTable.soap.setField("LinksEnum", list);
	if ( LinkTable.soap.parseResponse("<%= URL_WEB_SERVICES %>/LinksTo_Delete.asp<%= AspSiteQuery %>") )
	{
		LinkTable.soap.setField("IdObject", AspIdObject);
		LinkTable.query();
	}
}

function DoProtocolChange(select)
{
	var target = document.getElementById("FormLink_Link");
	var value = select.value;
	if (value != "")
		target.value = value + target.value.replace(/^\S*?:\/*/, "");
	if (value == "file://")
	{
		document.getElementById("FormLink_Label").innerText = "File:";
		document.getElementById("FormLink_Ellipses").style.visibility = "visible";
	}
	else
	{
		document.getElementById("FormLink_Label").innerText = "URL:";
		document.getElementById("FormLink_Ellipses").style.visibility = "hidden";
	}
}

//-->
</script>

		<TABLE CLASS="Dlg-clsTABLE" CELLSPACING=0 CELLPADDING=0 BORDER=0 STYLE="width: 100%; height: 100%;">
		<TR>
			<TD CLASS="ColorOutset" STYLE="height: 110px; padding: 2;">
			<FORM ID="FormLink" NAME="FormLink" STYLE="width: 100%; height: 100%; margin: 0; padding: 0; text-align: center;" onsubmit="return false;">
				<TABLE CLASS="ColorOutsetText" STYLE="width: 100%; xtable-layout: fixed;">
				<COL WIDTH=70><COL><COL WIDTH=25>
				<TR>
					<TD CLASS="KnoS-clsFont" ALIGN=right>
						<%= Language.label("Tipo di rif.") %>:
					</TD>
					<TD CLASS="KnoS-clsFont" STYLE="text-align: left; vertical-align: middle; padding-bottom: 4;">
						<INPUT TYPE="radio" ID="RadioLink_Interno" NAME="RadioLink" CHECKED STYLE="padding: 2;" onclick="DoRadioLinkInterno();"> <%= Language.label("Interno") %> </INPUT>
						<INPUT TYPE="radio" ID="RadioLink_Esterno" NAME="RadioLink" STYLE="padding: 2;" onclick="DoRadioLinkEsterno();"> <%= Language.label("Esterno") %> </INPUT>
					</TD>
					<TD WIDTH=25> &nbsp; </TD>
				</TR><TR ID="Protocollo_idTR" STYLE="display: none;">
					<TD CLASS="KnoS-clsFont" ALIGN=right>
						<%= Language.label("Protocollo") %>:
					</TD>
					<TD COLSPAN=2>
						<SELECT CLASS="KnoS-clsFont" ID="Protocol" SIZE="1" STYLE="width: 80; display: none;" onchange="DoProtocolChange(this);">
							<OPTION VALUE="" SELECTED>&nbsp;</OPTION>
							<OPTION VALUE="http://">http:</OPTION>
							<OPTION VALUE="https://">https:</OPTION>
							<OPTION VALUE="ftp://">ftp:</OPTION>
							<OPTION VALUE="file://">file:</OPTION>
							<OPTION VALUE="mailto:">mailto:</OPTION>
							<OPTION VALUE="news:">news:</OPTION>
						</SELECT>
					</TD>
				</TR><TR>
					<TD CLASS="KnoS-clsFont" ID="FormLink_Label" ALIGN=right>
						<%= Language.label("Id pubbl.") %>:
					</TD>
					<TD STYLE="padding-right: 0px;">
						<INPUT TYPE=text CLASS="KnoS-clsFont" NAME="FormLink_Link" ID="FormLink_Link" STYLE="width: 100%;" onpaste="DoLinkPaste();">
					</TD>
					<TD STYLE="width: 25px; text-align: left; padding-left: 0px;">
						<INPUT TYPE=button CLASS="clsINPUT-ellipses" NAME="Form_Ellipses" ID="FormLink_Ellipses" VALUE="..." HIDEFOCUS onclick="DoLinkEllipses('FormLink_Link');">
					</TD>
				</TR><TR>
					<TD CLASS="KnoS-clsFont" ALIGN=right>
						<%= Language.label("Descrizione") %>:
					</TD>
					<TD COLSPAN="2">
						<INPUT TYPE=text CLASS="KnoS-clsFont" NAME="FormLink_Descr" ID="FormLink_Descr" STYLE="width: 100%;">
					</TD>
				</TR><TR>
					<TD> &nbsp; </TD>
					<TD COLSPAN=2 ALIGN=right>
						<INPUT TYPE=button CLASS="KnoS-clsFont" NAME="Form_Save" ID="Form_Save" VALUE="<%= Language.label("Salva") %>" STYLE="width: 60;" HIDEFOCUS onclick="DoLinkSave();">
						<INPUT TYPE=button CLASS="KnoS-clsFont" NAME="Form_Cancel" ID="Form_Cancel" VALUE="<%= Language.label("Annulla") %>" STYLE="width: 60;" HIDEFOCUS onclick="DoLinkCancel();">
					</TD>
				</TR>
				</TABLE>
			</FORM>
			</TD>
		</TR>
		<TR>
			<TD CLASS="" STYLE="padding: 1;">
				<DIV CLASS="Dlg-clsDIV-table" ID="Link_idDIV" STYLE="background: window; overflow: auto; width: 100%;">
<script language="JavaScript">
<!--
	LinkTable.writeHTML();
//-->
</script>
				</DIV>
			</TD>
		</TR>
		</TABLE>
	</DIV>
</TD></TR>
</TABLE>

<!------------------------------- TAB LINK END ------------------------------->

<!--------------------------- TAB DESTINATARI START -------------------------->

<script language="JavaScript">
<!--

// INDICE ALFABETICO

var AZKey = AZKeyPad_2row("AZKey");

function AZKey_onclick(value)
{
	value = Trim(value.toLowerCase());
	SubjectTable.searchExpression = RadioFilter() + " AND "+ Subject_viewField +" >= '" + Trim(value) + "'";
	SubjectCursor.first();
};
AZKey.onclick = AZKey_onclick;


//-->
</script>


<TABLE ID="TabPermission_idTABLE" CLASS="CTabBar-clsTABLE-inner" CELLSPACING=0 CELLPADDING=0 STYLE="display: none;">
<TR><TD>
	<DIV STYLE="overflow: auto; width: 100%; height: 100%;">
	<TABLE CLASS="Main-clsTABLE" ID="Main_idTABLE" CELLPADDING=0 CELLSPACING=0>
	<COL CLASS="Left-clsCOL" ID="Left_idCOL" STYLE="width:50%;"><COL CLASS="Splitter-clsCOL" ID="Splitter_idCOL">
	<TR>

		<!-- LEFT -->

		<TD CLASS="Left-clsTD" ID="Left_idTD">
		<DIV CLASS="Left-clsDIV" ID="Left_idDIV" STYLE="padding:0; padding-left:0;">
			<TABLE CELLPADDING=0 CELLSPACING=0 STYLE="table-layout: fixed; width: 100%; height: 100%">
			<TR><TD CLASS="KnoS-clsFont ColorOutset" WIDTH=100% HEIGHT=40 NOWRAP VALIGN=top STYLE="padding: 0;">
<!-- Indice alfabetico -->
<script language="JavaScript">
<!--
	AZKey.writeHTML();
//-->
</script>
			</TD></TR>

			<!-- Ricerca testuale -->

			<TR><TD WIDTH=100% HEIGHT=67 NOWRAP VALIGN=top>
				<FORM CLASS="clsFORM" STYLE="margin:0;" onsubmit="RicercaTestuale(); return false;">
				<TABLE CELLPADDING=0 CELLSPACING=0 STYLE="width: 100%; margin: 0; border-width: 0;">
				<TR><TD CLASS="KnoS-clsFont ColorOutset" VALIGN=middle NOWRAP STYLE="padding: 0; padding-left: 4;" NOWRAP>
					<INPUT TYPE="radio" NAME="RadioUserView" ID="IdRadio_Subject" onclick="DoRadioUserViewChange();"> <%= Language.label("Nome") %> </INPUT>
					<INPUT TYPE="radio" NAME="RadioUserView" ID="IdRadio_FullName" onclick="DoRadioUserViewChange();"> <%= Language.label("Nome completo") %> </INPUT>
				</TD></TR>
				<TR><TD CLASS="ColorOutset" COLSPAN=2 STYLE="padding: 1; padding-bottom: 0; border-bottom-width: 0;" NOWRAP>
					<INPUT TYPE="text" CLASS="KnoS-clsFont clsTEXT_Ricerca" ID="idTEXT_Ricerca" STYLE="width:100%;">
				</TD></TR>
				<TR><TD CLASS="ColorOutset" COLSPAN=2 STYLE="padding:1; padding-top: 0; border-top-width: 0;">
					<TABLE CLASS="KnoS-clsFont ColorOutsetText" WIDTH=100% HEIGHT=20 CELLPADDING=0 CELLSPACING=0>
					<TR><TD ALIGN=left VALIGN=middle NOWRAP>
						<INPUT TYPE="radio" NAME="Radio" ID="IdRadio_Tutti" CHECKED onclick="DoRadioChange();"> <%= Language.label("Tutti") %> </INPUT>
						<INPUT TYPE="radio" NAME="Radio" ID="IdRadio_Gruppi" onclick="DoRadioChange();"> <%= Language.label("Gruppi") %> </INPUT>
						<INPUT TYPE="radio" NAME="Radio" ID="IdRadio_Utenti" onclick="DoRadioChange();"> <%= Language.label("Utenti") %> </INPUT>
						&nbsp;
					</TD><TD ALIGN=right>
						<INPUT ALIGN=right TYPE="submit" CLASS="KnoS-clsFont" NAME="idBUTTON_Ricerca" VALUE="<%= Language.label("Trova") %>" STYLE="width: 50;">
					</TD></TR>
					</TABLE>
				</TD></TR>
				</TABLE>
				</FORM>
			</TD></TR>

			<!-- Tabella Soggetti e cursore-->

			<TR><TD ID="Subject_idTD" CLASS="ColorOutsetEmpty" VALIGN=top NOWRAP>
				<DIV STYLE="padding: 1; width: 100%; height: 100%; overflow: auto; vertical-align: top;">
<script language="JavaScript">
<!--
	var SubjectRecord = new CRecord(["IdSubject", "FlagSubject", "Subject", "FullName", "FlagAuto"]);
	var SubjectTable = new CTable("SubjectTable", URL_WEB_SERVICES +"/Subjects_Items.asp"+ AspSiteQuery);
	SubjectTable.enableCheck = true;
	SubjectTable.view = ViewSubject();
	var SubjectCursor = new CCursorBar("SubjectCursor", SubjectTable);
	SubjectTable.writeHTML();
//-->
</script>
				</DIV>
			</TD></TR>
			<TR>
				<TD CLASS="ColorOutset" STYLE="height: 25; padding: 0; border-width: 0;">
<script language="JavaScript">
<!--
	SubjectCursor.writeHTML();
//-->
</script>
				</TD>
			</TR>
			</TABLE>
		</DIV>
		</TD>

		<!-- SPLITTER -->

		<TD CLASS="HSplitter-clsTD" ID="Splitter_idTD">
			<DIV CLASS="HSplitter-clsDIV-normal" ID="Splitter_idDIV" STYLE="cursor: default;">
			</DIV>
		</TD>

		<!-- RIGHT -->

		<TD CLASS="Right-clsTD" ID="Right_idTD">
			<DIV CLASS="Right-clsDIV" ID="Right_idDIV">
				<TABLE CLASS="ColorInset" CELLPADDING=0 CELLSPACING=0 STYLE="table-layout: fixed; width: 100%; height: 100%">

				<!-- Tabella Soggetti selezionati -->

				<TR><TD CLASS="ColorOutsetEmpty" VALIGN=top NOWRAP>
					<DIV STYLE="padding: 1; width: 100%; height: 100%; overflow: auto; vertical-align: top;">
<script language="JavaScript">
<!--
	var SelectedTable = new CTable("SelectedTable", URL_WEB_SERVICES +"/Object_GetPermissions.asp"+ AspSiteQuery);
	SelectedTable.pageSize = 0;
	SelectedTable.enableCheck = true;
	SelectedTable.view = ViewSubject();
	SelectedTable.writeHTML();
//-->
</script>
					</DIV>
				</TD></TR>
				<TR>
					<TD CLASS="ColorOutset" STYLE="height: 24; padding: 1; text-align: center;" NOWRAP>

					<!-- Bottoni -->
							<INPUT TYPE=button CLASS="KnoS-clsFont" VALUE="<%= Language.label("Cancella") %>" HIDEFOCUS STYLE="width: 60;" onclick="DoEliminaSubject();"></INPUT>
					</TD>
				</TR>
				</TABLE>
			</DIV>
		</TD>
	</TR>
	</TABLE>
	</DIV>
</TD></TR>
</TABLE>

<!---------------------------- TAB DESTINATARI END --------------------------->

<!------------------------------ TAB AZIONI START ---------------------------->

<TABLE ID="TabAction_idTABLE" CLASS="CTabBar-clsTABLE-inner" CELLSPACING=0 CELLPADDING=0 STYLE="display: none;">
<TR><TD>
	<DIV STYLE="width: 100%; height: 100%;">
	<TABLE CLASS="Dlg-clsTABLE" CELLSPACING=0 CELLPADDING=0 BORDER=0 STYLE="width: 100%; height: 100%; table-layout: fixed;">
	<TR>
		<TD STYLE="vertical-align: top;" NOWRAP>
			<DIV STYLE="width: 100%; height: 100%; overflow: auto;">

<script language="JavaScript">
<!--

var ActionTable = new CTable("ActionTable", "<%= URL_WEB_SERVICES %>/Object_Actions.asp<%= AspSiteQuery %>");
ActionTable.caption = Language.label("Selezionare l'azione da eseguire");
ActionTable.pageSize = 0;
ActionTable.view = ViewAction();
ActionTable.autoselect = true;
ActionTable.ondblclick = function (tr) { ActionTable.select(tr); DoAction(); };
ActionTable.writeHTML();

//-->
</script>
			</DIV>
		</TD>
	</TR>
	<TR>
		<TD CLASS="Dlg-clsTD-bottom ColorOutset" ALIGN=center STYLE="padding: 2; height: 26px;">
			<INPUT TYPE=button CLASS="KnoS-clsFont" ID="Refresh_idBUTTON" VALUE="<%= Language.label("Esegui azione") %>" STYLE="width: 90;" HIDEFOCUS onclick="DoAction();">
		</TD>
	</TR>
	</TABLE>
	</DIV>
</TD></TR>
</TABLE>

<!------------------------------- TAB AZIONI END ----------------------------->


<!------------------------------ TAB WORKFLOW START ---------------------------->

<TABLE ID="TabWorkflow_idTABLE" CLASS="CTabBar-clsTABLE-inner" CELLSPACING=0 CELLPADDING=0 STYLE="display: none;">
<TR><TD>
	<DIV STYLE="width: 100%; height: 100%; overflow: auto;">
	<script language="JavaScript">
	<!--
		var Workflow = template = new CWorkflow("Workflow");
		Workflow.writeHTML();
	//-->
	</script>
	</DIV>
</TD></TR>
</TABLE>

<!------------------------------- TAB WORKFLOW END ----------------------------->


<!----------------------------- TAB CATALOGHI START -------------------------->

<TABLE ID="TabCatalog_idTABLE" CLASS="CTabBar-clsTABLE-inner" CELLSPACING=0 CELLPADDING=0 STYLE="display: none;">
<TR>
	<TD STYLE="height: 24px;">
	<TABLE CLASS="KnoS-clsFont ColorOutset" CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD STYLE="padding-left: 4px; padding-right: 4px; text-align: right; width: 70px;"><%= Language.label("Catalogo") %>:</TD>
		<TD>
			<table cellspacing="0" cellpadding="0">
			<col /><col style="width: 100px"/>
			<tr>
				<td>
				<div ID="TabCatalog_idDIV_Catalog">
					<script language="JavaScript">
					<!--
					var CatalogTree = new CTreeView("CatalogTree");
					CatalogTree.checked = true;
					CatalogTree.value.IdTree = 0;
					CatalogTree.oncheck = function(node, checkbox, nRow)
					{
						var selLink = ComboLinkType.getValue();
						if (node.LinkType == null)
							node.LinkType = [];
						Soap.setField("IdTree", CatalogTree.value.IdTree);
						Soap.setField("IdNode", node.id);
						Soap.setField("IdObject", AspIdObject);
						Soap.setField("IdAttr", 0);
						Soap.setField("LinkValue", selLink.LinkValue);
						Soap.setField("LinkType", selLink.LinkType);
						if (checkbox.checked)
						{
							if (Soap.parseResponse(URL_WEB_SERVICES +"/Catalog_AddRef.asp"+ AspSiteQuery))
								node.LinkType[selLink.LinkType] = selLink.LinkValue;
							else
								checkbox.checked = false;
						}
						else
						{
							if (Soap.parseResponse(URL_WEB_SERVICES +"/Catalog_DeleteRef.asp"+ AspSiteQuery))
								try {delete(node.LinkType[selLink.LinkType])}catch(e){};
							else
								checkbox.checked = true;
						}
					}

					var ComboCatalog = new CComboBox("ComboCatalog");
					ComboCatalog.colHidden = 1;
					ComboCatalog.colDisplay = true;
					ComboCatalog.colSelected = 1;
					ComboCatalog.col = ["IdTree", Language.label("Catalogo"), Language.label("Collegamenti")];
					ComboCatalog.row = [];
					ComboCatalog.listSize = 8;
					ComboCatalog.writeHTML();
					ComboCatalog.onclick = function(nRow)
					{
						var idTree = this.row[nRow][0];
						ReloadCatalogTree(idTree);
					}

					//-->
					</script>
				</div>
				</td>
				<td>
					<input type="button" value="<%= Language.label("Aggiorna") %>" style="width: 100%;" onclick="ReloadComboCatalog(CatalogTree.value.IdTree)"/>
				</td>
			</tr>
			</table>
		</TD>
	</TR>
	<TR>
		<TD STYLE="padding-left: 4px; padding-right: 4px; text-align: right;" NOWRAP><%= Language.label("col: ValueDescr") %>:</TD>
		<TD STYLE="padding: 0px; height: 24px;">
		<DIV ID="TabCatalog_idDIV_LinkType">
			<% AspTemplate = "ComboLinkType" %>
			<!-- #include virtual="/knos/web/incCatalogLinkTypeComboBox.asp" -->		
			<script language="JavaScript">
			<!--
				ComboLinkType.label = "";
				ComboLinkType.tableStyle = "width: 300px;";
				ComboLinkType.writeHTML();
				ComboLinkType.sync(AspLinkTypeList.Elemento.LinkType);
				ComboLinkType.onclick = function()
				{
					RefreshCatalogTree();
				}
			//-->
			</script>
		</DIV>
		</TD>
	</TR>
	</TABLE>
	</TD>
</TR>
</TR>
<TR>
	<TD>
	<DIV CLASS="ColorInset" STYLE="width: 100%; height: 100%; text-align: left; overflow: auto; background-color: white;" ID="FolderTree_idDIV">
		&nbsp;
	</DIV>
	</TD>
</TR>
</TABLE>

<!------------------------------ TAB CATALOGHI END --------------------------->


<script language="JavaScript">
<!--
TabBar.writeHTML_END();
//-->
</script>

		</TD>
	</TR>
	<TR>
		<TD ID="ButtonList_TD_Enabled" CLASS="Dlg-clsTD-bottom ColorOutset" COLSPAN=2 ALIGN=center STYLE="padding: 2;" style="display:none">
			<INPUT TYPE=button name="ObjectFullEdit_BUTTON" CLASS="KnoS-clsFont" ID="Save_idBUTTON" VALUE="<%= Language.label("Salva") %>" STYLE="width: 80; display: none;" HIDEFOCUS onclick="DoSalva();">
			<INPUT TYPE=button name="ObjectFullEdit_BUTTON" CLASS="KnoS-clsFont" ID="SaveExit_idBUTTON" VALUE="<%= Language.label("Salva ed esci") %>" STYLE="width: 80; display: none;" HIDEFOCUS onclick="DoSalva(true);">
			<INPUT TYPE=button name="ObjectFullEdit_BUTTON" CLASS="KnoS-clsFont" ID="Notify_idBUTTON" VALUE="<%= Language.label("Notifica") %>" STYLE="width: 80; display: none;" HIDEFOCUS onclick="DoNotify();">
			<INPUT TYPE=button name="ObjectFullEdit_BUTTON" CLASS="KnoS-clsFont" ID="Cancel_idBUTTON" VALUE="<%= Language.label("Annulla") %>" STYLE="width: 80; display: none;" HIDEFOCUS onclick="DoAnnulla();">
			<INPUT TYPE=button name="ObjectFullEdit_BUTTON" CLASS="KnoS-clsFont" ID="Close_idBUTTON" VALUE="<%= Language.label("Chiudi") %>" STYLE="width: 80; display: none;" HIDEFOCUS onclick="DoChiudi();">
			<INPUT TYPE=button name="ObjectFullEdit_BUTTON" CLASS="KnoS-clsFont" ID="Next_idBUTTON" VALUE="<%= Language.label("Avanti") %>" STYLE="width: 80; display: none;" HIDEFOCUS onclick="DoAvanti();">
			<INPUT TYPE=button name="ObjectFullEdit_BUTTON" CLASS="KnoS-clsFont" ID="Undo_idBUTTON" VALUE="<%= Language.label("Rinuncia") %>" STYLE="width: 80; display: none;" HIDEFOCUS onclick="DoRinuncia(true);">
		</TD>
	</TR>
	</TABLE>

<script src="/knos/web/DlgObjectFullEdit.js?<%= ANTICACHE_BUILD %>"></script>
<script language="JavaScript">

function DoObjectCheckIn()
{
	Soap.setField("IdObject", AspIdObject);
	if (Soap.parseResponse("<%= URL_WEB_SERVICES %>/Object_CheckIn.asp<%= AspSiteQuery %>"))
		return true;

	return false;
}

 ////////
// Exit
//
// return true se esegue la window.close()
//
function Exit(result)
{
	try
	{
	    if (result == null)
			result = 0;
		switch (result)
		{
			case 4:	// Rinuncia: si elimina la pubblicazione 
				Soap.setField("IdObject", AspIdObject);
				Soap.parseResponse(URL_WEB_SERVICES+"/Object_Delete.asp"+ AspSiteQuery);
				break;
			case 5: // Avanti: si applica il template se specificato
				if (AspAction == "matchtemplate")
				{
					Soap.setField("IdObject", AspIdObject);
					Soap.setField("MergeWithoutTemplateAttribute", "<Vs/>");
					Soap.setField("MergeTemplateStatus", 1);
					if (Soap.parseResponse(URL_WEB_SERVICES+"/Object_SaveEx.asp"+ AspSiteQuery) )
						document.location = document.location.href.replace(/Action=MatchTemplate/i, "Action=");
					return false;
				}
				break;
			default:
				DoObjectCheckIn();
		}
		window.onunload = EmptyFunction;
		window.external.result = result;
		window.external.Exit();
		return true;
	}
	catch(e){}

	try
	{
		if (window.opener != null || window.external.returnValue != null)
		{
			window.close();
			return true;
		}
	}
	catch(e)
	{ }

	return false;
}

</script>

<!-- #include virtual="/custom/incCustomBottom.asp" -->
</BODY>
</HTML>