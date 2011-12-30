<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" --> 
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");	
	var rsServiceProvider = Server.CreateObject("ADODB.Recordset");
	rsServiceProvider.ActiveConnection = MM_cnnASP02_STRING;
	rsServiceProvider.Source = "{call dbo.cp_srv_pvdr(0,'"+Description+"',0,'A',0)}";
	rsServiceProvider.CursorType = 0;
	rsServiceProvider.CursorLocation = 2;
	rsServiceProvider.LockType = 3;
	rsServiceProvider.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");	
}
%>
<html>
<head>
	<title>New Service Provider</title>
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
		if (Trim(document.frm0398.Description.value)=="") {
			alert("Enter Description.");
			document.frm0398.Description.focus();
			return ;
		}
		document.frm0398.submit();
	}
	</script>
</head>
<body onLoad="document.frm0398.Description.focus();">
<form name="frm0398" method="POST" action="<%=MM_editAction%>">
<h5>New Service Provider</h5>
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