<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_edit"))=="true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var rsInstitutionType = Server.CreateObject("ADODB.Recordset");
	rsInstitutionType.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitutionType.Source = "{call dbo.cp_school_type("+Request.QueryString("insSchool_type_id")+",'" + Description + "'," + IsActive + ",0,'E',0)}";
	rsInstitutionType.CursorType = 0;
	rsInstitutionType.CursorLocation = 2;
	rsInstitutionType.LockType = 3;
	rsInstitutionType.Open();
	Response.Redirect("m018q0321.asp");
}

var rsInstitutionType = Server.CreateObject("ADODB.Recordset");
rsInstitutionType.ActiveConnection = MM_cnnASP02_STRING;
rsInstitutionType.Source = "{call dbo.cp_school_type("+Request.QueryString("insSchool_type_id")+",'',0,1,'Q',0)}";
rsInstitutionType.CursorType = 0;
rsInstitutionType.CursorLocation = 2;
rsInstitutionType.LockType = 3;
rsInstitutionType.Open();
%>
<html>
<head>
	<title>Update Institution Type Lookup</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0321.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Save(){
		if (Trim(document.frm0321.Description.value)==""){
			alert("Enter Description.");
			document.frm0321.Description.focus();
			return ;
		}
		document.frm0321.submit();
	}
	</script>
</head>
<body onLoad="document.frm0321.Description.focus();">
<form name="frm0321" method="POST" action="<%=MM_editAction%>">
<h5>Update Institution Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="50" size="30" tabindex="1" accesskey="F" value="<%=(rsInstitutionType.Fields.Item("chvSchool_Type").Value)%>"></td>
	</tr>
	<tr>
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" tabindex="2" value="1" accesskey="L" <%=((rsInstitutionType.Fields.Item("bitIs_active").Value=="1")?"CHECKED":"")%> class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="3" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="5" onClick="history.back()" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_edit" value="true">
</form>
</body>
</html>
<%
rsInstitutionType.Close();
%>