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
	var IsActive = ((Request.Form("IsActive")=="on") ? "1":"0");
	var IsManSelect = ((Request.Form("IsManualSelect")=="on") ? "1":"0");
	var rsInventoryStatus = Server.CreateObject("ADODB.Recordset");
	rsInventoryStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsInventoryStatus.Source = "{call dbo.cp_Insert_Equip_Status('"+ Description + "'," + IsActive + "," + IsManSelect + ",0)}";
	rsInventoryStatus.CursorType = 0;
	rsInventoryStatus.CursorLocation = 2;
	rsInventoryStatus.LockType = 3;
	rsInventoryStatus.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Inventory Status Lookup</title>
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
				document.frm0336.reset();
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
		if (Trim(document.frm0336.Description.value)==""){
			alert("Enter Description.");
			document.frm0336.Description.focus();
			return ;
		}
		document.frm0336.submit();
	}
	</script>
</head>
<body onLoad="document.frm0336.Description.focus();">
<form name="frm0336" method="POST" action="<%=MM_editAction%>">
<h5>New Inventory Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr>
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" tabindex="2" class="chkstyle"></td>
	</tr>
    <tr>
		<td>Is Manual Select:</td>
		<td><input type="checkbox" name="IsManualSelect" tabindex="3" accesskey="L" class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="4" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="5" onClick="window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>