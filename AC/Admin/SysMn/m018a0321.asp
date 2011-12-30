<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}
if (String(Request("MM_Insert")) != "undefined") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var IsActive = ((String(Request.Form("IsActive"))=="1")?"1":"0");
	var rsInstitutionType = Server.CreateObject("ADODB.Recordset");
	rsInstitutionType.ActiveConnection = MM_cnnASP02_STRING;
	rsInstitutionType.Source = "{call dbo.cp_school_type(0,'"+ Description + "'," + IsActive + ",0,'A',0)}";
	rsInstitutionType.CursorType = 0;
	rsInstitutionType.CursorLocation = 2;
	rsInstitutionType.LockType = 3;
	rsInstitutionType.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Institution Type</title>
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
			case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Save(){
		if (Trim(document.frm0321.Description.value)=="") {
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
<h5>New Institution Type</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
	</tr>
	<tr>
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" value="1" tabindex="2" class="chkstyle" accesskey="L"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="4" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
