<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_Insert")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var IsBuyout = ((Request.Form("IsBuyout")=="1") ? "1":"0");	
	var rsEquipUserType = Server.CreateObject("ADODB.Recordset");
	rsEquipUserType.ActiveConnection = MM_cnnASP02_STRING;
	rsEquipUserType.Source = "{call dbo.cp_eq_user_type2(0,'"+Description+"',"+ IsActive + ","+ IsBuyout +",0,'A',0)}";
	rsEquipUserType.CursorType = 0;
	rsEquipUserType.CursorLocation = 2;
	rsEquipUserType.LockType = 3;
	rsEquipUserType.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Equipment User Type</title>
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
		if (Trim(document.frm03122.Description.value)==""){
			alert("Enter Description.");
			document.frm03122.Description.focus();
			return ;
		}
		document.frm03122.submit();
	}
	</script>
</head>
<body onLoad="document.frm03122.Description.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm03122">
<h5>New Equipment User Type</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="50" size="30" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr>
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" value="1" tabindex="2" class="chkstyle"></td>
	</tr>
    <tr>
		<td>Is Buyout:</td>
		<td><input type="checkbox" name="IsBuyout" value="1" tabindex="3" accesskey="L" class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="5" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>