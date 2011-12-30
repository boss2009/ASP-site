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
	var IsActive = ((Request.Form("IsActive")=="on") ? "1":"0");
	var rsProgramType = Server.CreateObject("ADODB.Recordset");
	rsProgramType.ActiveConnection = MM_cnnASP02_STRING;
	rsProgramType.Source = "{call dbo.cp_program_type2(0,'"+ Description + "'," + IsActive + ",0,'A',0)}";
	rsProgramType.CursorType = 0;
	rsProgramType.CursorLocation = 2;
	rsProgramType.LockType = 3;
	rsProgramType.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Program Type</title>
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
		if (Trim(document.frm0320.Description.value)=="") {
			alert("Enter Description.");
			document.frm0320.Description.focus();
			return ;
		}
		document.frm0320.submit();
	}
	</script>
</head>
<body onLoad="document.frm0320.Description.focus();">
<form name="frm0320" method="POST" action="<%=MM_editAction%>">
<h5>New Program Type</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
	</tr>
	<tr>
		<td>IsActive:</td>
		<td><input type="checkbox" name="IsActive" tabindex="2" accesskey="L" class="chkstyle"></td>
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