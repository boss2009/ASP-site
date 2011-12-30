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
	var rsPILATStatus = Server.CreateObject("ADODB.Recordset");
	rsPILATStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsPILATStatus.Source = "{call dbo.cp_pilat_status("+Request.QueryString("insPILAT_Status_id")+",'"+Description+"',0,'E',0)}";
	rsPILATStatus.CursorType = 0;
	rsPILATStatus.CursorLocation = 2;
	rsPILATStatus.LockType = 3;
	rsPILATStatus.Open();
	Response.Redirect("m018q03109.asp");
}

var rsPILATStatus = Server.CreateObject("ADODB.Recordset");
rsPILATStatus.ActiveConnection = MM_cnnASP02_STRING;
rsPILATStatus.Source = "{call dbo.cp_pilat_status("+Request.QueryString("insPILAT_Status_id")+",'',1,'Q',0)}";
rsPILATStatus.CursorType = 0;
rsPILATStatus.CursorLocation = 2;
rsPILATStatus.LockType = 3;
rsPILATStatus.Open();
%>
<html>
<head>
	<title>Update PILAT Status Lookup</title>
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
				document.frm03109.reset();
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
		if (Trim(document.frm03109.Description.value)==""){
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
<h5>Update PILAT Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="<%=(rsPILATStatus.Fields.Item("chvStatus_Desc").Value)%>" maxlength="40" size="40" tabindex="1" accesskey="F" ></td>
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
<input type="hidden" name="MM_recordId" value="<%= rsPILATStatus.Fields.Item("insPILAT_Status_id").Value %>">
</form>
</body>
</html>
<%
rsPILATStatus.Close();
%>