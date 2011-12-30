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
	var rsShippingStatus = Server.CreateObject("ADODB.Recordset");
	rsShippingStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsShippingStatus.Source = "{call dbo.cp_ship_rtn_status(0,'"+Description+"',0,'A',0)}";
	rsShippingStatus.CursorType = 0;
	rsShippingStatus.CursorLocation = 2;
	rsShippingStatus.LockType = 3;
	rsShippingStatus.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Shipping Status</title>
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
		if (Trim(document.frm03152.Description.value)=="") {
			alert("Enter Description.");
			document.frm03152.Description.focus();
			return ;
		}
		document.frm03152.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03152.Description.focus();">
<form name="frm03152" method="POST" action="<%=MM_editAction%>">
<h5>New Shipping Status</h5>
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
		<td><input type="button" value="Save" onClick="Save();" tabindex="2" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="3" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>