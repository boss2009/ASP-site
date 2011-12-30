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
	var rsTrainingStatus = Server.CreateObject("ADODB.Recordset");
	rsTrainingStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsTrainingStatus.Source = "{call dbo.cp_Training_status(0,'"+Description+"',0,'A',0)}";
	rsTrainingStatus.CursorType = 0;
	rsTrainingStatus.CursorLocation = 2;
	rsTrainingStatus.LockType = 3;
	rsTrainingStatus.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Training Status</title>
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
		if (Trim(document.frm03137.Description.value)=="") {
			alert("Enter Description.");
			document.frm03137.Description.focus();
			return ;
		}
		document.frm03137.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03137.Description.focus();">
<form name="frm03137" method="POST" action="<%=MM_editAction%>">
<h5>New Training Status</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F" ></td>
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