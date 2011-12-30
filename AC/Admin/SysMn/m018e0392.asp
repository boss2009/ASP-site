<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}
if (String(Request("MM_update")) == "true") {	
	var Description = String(Request.Form("Description")).replace(/'/g, "''");	
	var rsRepairStatus = Server.CreateObject("ADODB.Recordset");
	rsRepairStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsRepairStatus.Source = "{call dbo.cp_repair_status("+ Request.Form("MM_recordId") + ",'" + Description + "',0,'E',0)}";
	rsRepairStatus.CursorType = 0;
	rsRepairStatus.CursorLocation = 2;
	rsRepairStatus.LockType = 3;
	rsRepairStatus.Open();
	Response.Redirect("m018q0392.asp");
}

var rsRepairStatus = Server.CreateObject("ADODB.Recordset");
rsRepairStatus.ActiveConnection = MM_cnnASP02_STRING;
rsRepairStatus.Source = "{call dbo.cp_repair_status("+ Request.QueryString("insEq_Repair_Sts_id") + ",'',1,'Q',0)}";
rsRepairStatus.CursorType = 0;
rsRepairStatus.CursorLocation = 2;
rsRepairStatus.LockType = 3;
rsRepairStatus.Open();
%>
<html>
<head>
	<title>Update Repair Status Lookup</title>
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
<h5>Update Repair Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="<%=(rsRepairStatus.Fields.Item("chvEq_Repair_Sts_Desc").Value)%>" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
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
<input type="hidden" name="MM_recordId" value="<%=rsRepairStatus.Fields.Item("insEq_Repair_Sts_id").Value %>">
</form>
</body>
</html>
<%
rsRepairStatus.Close();
%>