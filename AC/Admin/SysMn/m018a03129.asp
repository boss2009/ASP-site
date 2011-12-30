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
	var AreaCode = String(Request.Form("AreaCode")).replace(/'/g, "''");			
	var IsLocal = ((Request.Form("IsLocal")=="1") ? "1":"0");
	var rsAreaCode = Server.CreateObject("ADODB.Recordset");
	rsAreaCode.ActiveConnection = MM_cnnASP02_STRING;
	rsAreaCode.Source = "{call dbo.cp_area_code(0,'"+ AreaCode + "'," + IsLocal + ",0,'A',0)}";
	rsAreaCode.CursorType = 0;
	rsAreaCode.CursorLocation = 2;
	rsAreaCode.LockType = 3;
	rsAreaCode.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Area Code Lookup</title>
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
				document.frm03129.reset();
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
		if (Trim(document.frm03129.AreaCode.value)==""){
			alert("Enter Area Code.");
			document.frm03129.AreaCode.focus();
			return ;		
		}
		document.frm03129.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03129.AreaCode.focus();">
<form name="frm03129" method="POST" action="<%=MM_editAction%>">
<h5>New Area Code Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Area Code:</td>
		<td nowrap><input type="text" name="AreaCode" maxlength="3" size="3" onKeypress="AllowNumericOnly();" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Is Local:</td>
		<td nowrap><input type="checkbox" name="IsLocal" value="1" class="chkstyle" tabindex="2" accesskey="L"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="3" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="4" onClick="window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>