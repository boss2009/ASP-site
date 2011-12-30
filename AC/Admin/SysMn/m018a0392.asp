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
	var rsBuyoutStatus = Server.CreateObject("ADODB.Recordset");
	rsBuyoutStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsBuyoutStatus.Source = "{call dbo.cp_repair_status(0,'"+Description+"',0,'A',0)}";
	rsBuyoutStatus.CursorType = 0;
	rsBuyoutStatus.CursorLocation = 2;
	rsBuyoutStatus.LockType = 3;
	rsBuyoutStatus.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Repair Status Lookup</title>
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
				document.frm0392.reset();
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
		if (Trim(document.frm0392.Description.value)==""){
			alert("Enter Description.");
			document.frm0392.Description.focus();
			return ;		
		}
		document.frm0392.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0392.Description.focus();">
<form name="frm0392" method="POST" action="<%=MM_editAction%>">
<h5>New Repair Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="3" onClick="window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>