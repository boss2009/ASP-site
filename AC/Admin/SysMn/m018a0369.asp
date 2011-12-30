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
	var Abbreviation = String(Request.Form("Abbreviation")).replace(/'/g, "''");					
	var rsDuration = Server.CreateObject("ADODB.Recordset");
	rsDuration.ActiveConnection = MM_cnnASP02_STRING;
	rsDuration.Source = "{call dbo.cp_duratn_type2(0,'"+Description+"','"+ Abbreviation + "',0,'A',0)}";
	rsDuration.CursorType = 0;
	rsDuration.CursorLocation = 2;
	rsDuration.LockType = 3;
	rsDuration.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Duration Type Lookup</title>
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
				document.frm0369.reset();
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
		if (Trim(document.frm0369.Description.value)==""){
			alert("Enter Description.");
			document.frm0369.Description.focus();
			return ;		
		}
		document.frm0369.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0369.Description.focus();">
<form name="frm0369" method="POST" action="<%=MM_editAction%>">
<h5>New Duration Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Abbreviation:</td>
		<td nowrap><input type="text" name="Abbreviation" maxlength="10" size="10" tabindex="2" accesskey="L"></td>
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