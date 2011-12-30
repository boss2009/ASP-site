<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}
if (String(Request("MM_Insert")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var rsWorkOrder = Server.CreateObject("ADODB.Recordset");
	var IsActive = ((Request.Form("IsActive")=="on") ? "1":"0");
	rsWorkOrder.ActiveConnection = MM_cnnASP02_STRING;
	rsWorkOrder.Source = "{call dbo.cp_Work_Order(0,'"+ Request.Form("WorkOrderNumber") + "','" + Description + "'," + IsActive + ",0,'A',0)}";
	rsWorkOrder.CursorType = 0;
	rsWorkOrder.CursorLocation = 2;
	rsWorkOrder.LockType = 3;
	rsWorkOrder.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Work Order</title>
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
		if (Trim(document.frm0359.Description.value)=="") {
			alert("Enter Description.");
			document.frm0359.Description.focus();
			return ;
		}
		document.frm0359.submit();
	}
	</script>
</head>
<body onLoad="document.frm0359.Description.focus();">
<form name="frm0359" method="POST" action="<%=MM_editAction%>">
<h5>New Work Order</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
	</tr>
	<tr>
		<td nowrap>Work Order Number:</td>
		<td nowrap><input type="text" name="WorkOrderNumber" maxlength="40" size="20" tabindex="2"></td>
	</tr>
	<tr>
		<td nowrap>Is Active:</td>
		<td nowrap><input type="checkbox" name="IsActive" tabindex="3" accesskey="L" class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="5" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>