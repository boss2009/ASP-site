<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
// set the form action variable
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}
if (String(Request("MM_Insert")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var IsActive = ((Request.Form("IsActive")=="on") ? "1":"0");
	var rsPurchaseStatus = Server.CreateObject("ADODB.Recordset");
	rsPurchaseStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsPurchaseStatus.Source = "{call dbo.cp_purchase_status(0,'"+ Description + "'," + IsActive + ",'A',0,0)}";
	rsPurchaseStatus.CursorType = 0;
	rsPurchaseStatus.CursorLocation = 2;
	rsPurchaseStatus.LockType = 3;
	rsPurchaseStatus.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Purchase Status</title>
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
		if (Trim(document.frm0354.Description.value)=="") {
			alert("Enter Description.");
			document.frm0354.Description.focus();
			return ;
		}
		document.frm0354.submit();
	}
	</script>
</head>
<body onLoad="document.frm0354.Description.focus();">
<form name="frm0354" method="POST" action="<%=MM_editAction%>">
<h5>New Purchase Status</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
	</tr>
	<tr>
		<td>Is Active:</td>
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