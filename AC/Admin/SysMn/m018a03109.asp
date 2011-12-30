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
	var rsPILATStatus = Server.CreateObject("ADODB.Recordset");
	rsPILATStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsPILATStatus.Source = "{call dbo.cp_pilat_status(0,'"+Description+"',0,'A',0)}";
	rsPILATStatus.CursorType = 0;
	rsPILATStatus.CursorLocation = 2;
	rsPILATStatus.LockType = 3;
	rsPILATStatus.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New PILAT Status Lookup</title>
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
		if (Trim(document.frm03109.Description.value)=="") {
			alert("Enter Description.");
			document.frm03109.Description.focus();
			return ;
		}
		document.frm03109.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03109.Description.focus();">
<form name="frm03109" method="POST" action="<%=MM_editAction%>">
<h5>New PILAT Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="20" tabindex="1" accesskey="F"></td>
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