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
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var rsMailList = Server.CreateObject("ADODB.Recordset");
	rsMailList.ActiveConnection = MM_cnnASP02_STRING;
	rsMailList.Source = "{call dbo.cp_Mail_List(0,'"+ Description + "'," + IsActive + ",0,'A',0)}";
	rsMailList.CursorType = 0;
	rsMailList.CursorLocation = 2;
	rsMailList.LockType = 3;
	rsMailList.Open();
	rsMailList.Close();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Mail List Lookup</title>
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
				document.frm03124.reset();
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
		if (Trim(document.frm03124.Description.value)==""){
			alert("Enter Description.");
			document.frm03124.Description.focus();
			return ;		
		}
		document.frm03124.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03124.Description.focus();">
<form name="frm03124" method="POST" action="<%=MM_editAction%>">
<h5>New Mail List Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="50" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" value="1" tabindex="2" accesskey="L" class="chkstyle"></td>
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