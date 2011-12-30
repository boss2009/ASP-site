<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var rsFundingSource = Server.CreateObject("ADODB.Recordset");
	rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
	rsFundingSource.Source = "{call dbo.cp_Funding_Source3(0,'" + Description + "'," + IsActive + ",0,'A',0)}";
	rsFundingSource.CursorType = 0;
	rsFundingSource.CursorLocation = 2;
	rsFundingSource.LockType = 3;
	rsFundingSource.Open();
	Response.Redirect("m018q0328.asp");
}
%>
<html>
<head>
	<title>New Funding Source</title>
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
		if (Trim(document.frm0328.Description.value)=="") {
			alert("Enter Description.");
			document.frm0328.Description.focus();
			return ;
		}
		document.frm0328.submit();
	}
	</script>
</head>
<body onLoad="document.frm0328.Description.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0328">
<h5>New Funding Source</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr>
		<td nowrap>Is Active:</td>
		<td nowrap><input type="checkbox" name="IsActive" value="1" tabindex="2" class="chkstyle" accesskey="L"></td>
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