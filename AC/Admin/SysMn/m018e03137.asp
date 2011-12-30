<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var rsTrainingStatus = Server.CreateObject("ADODB.Recordset");
	rsTrainingStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsTrainingStatus.Source = "{call dbo.cp_training_status("+Request.Form("MM_recordId")+",'"+Description+"',0,'E',0)}";
	rsTrainingStatus.CursorType = 0;
	rsTrainingStatus.CursorLocation = 2;
	rsTrainingStatus.LockType = 3;
	rsTrainingStatus.Open();
	Response.Redirect("m018q03137.asp");
}

var rsTrainingStatus = Server.CreateObject("ADODB.Recordset");
rsTrainingStatus.ActiveConnection = MM_cnnASP02_STRING;
rsTrainingStatus.Source = "{call dbo.cp_training_status("+ Request.QueryString("insTrain_Status_id") + ",'',1,'Q',0)}";
rsTrainingStatus.CursorType = 0;
rsTrainingStatus.CursorLocation = 2;
rsTrainingStatus.LockType = 3;
rsTrainingStatus.Open();
%>
<html>
<head>
	<title>Update Training Status Lookup</title>
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
				document.frm03137.reset();
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
		if (Trim(document.frm03137.Description.value)==""){
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
<h5>Update Training Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="<%=(rsTrainingStatus.Fields.Item("chvTrain_Status").Value)%>" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="4" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsTrainingStatus.Fields.Item("insTrain_Status_id").Value %>">
</form>
</body>
</html>
<%
rsTrainingStatus.Close();
%>