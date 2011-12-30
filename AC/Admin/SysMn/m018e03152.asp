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
	var rsShippingStatus = Server.CreateObject("ADODB.Recordset");
	rsShippingStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsShippingStatus.Source = "{call dbo.cp_ship_rtn_status("+Request.Form("MM_recordId")+",'"+Description+"',0,'E',0)}";
	rsShippingStatus.CursorType = 0;
	rsShippingStatus.CursorLocation = 2;
	rsShippingStatus.LockType = 3;
	rsShippingStatus.Open();
	Response.Redirect("m018q03152.asp");
}

var rsShippingStatus = Server.CreateObject("ADODB.Recordset");
rsShippingStatus.ActiveConnection = MM_cnnASP02_STRING;
rsShippingStatus.Source = "{call dbo.cp_ship_rtn_status("+ Request.QueryString("insRtn_to_User") + ",'',1,'Q',0)}";
rsShippingStatus.CursorType = 0;
rsShippingStatus.CursorLocation = 2;
rsShippingStatus.LockType = 3;
rsShippingStatus.Open();
%>
<html>
<head>
	<title>Update Shipping Status Lookup</title>
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
			document.frm03152.reset();
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
		if (Trim(document.frm03152.Description.value)==""){
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
<h5>Update Shipping Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsShippingStatus.Fields.Item("chvRtoUser_Desc").Value)%>" maxlength="40" size="20" tabindex="1" accesskey="F" ></td>
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
<input type="hidden" name="MM_recordId" value="<%= rsShippingStatus.Fields.Item("insRtn_to_User").Value %>">
</form>
</body>
</html>
<%
rsShippingStatus.Close();
%>