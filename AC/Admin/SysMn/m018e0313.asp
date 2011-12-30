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
	var Abbreviation = String(Request.Form("Abbreviation")).replace(/'/g, "''");	
	var rsPhoneType = Server.CreateObject("ADODB.Recordset");
	rsPhoneType.ActiveConnection = MM_cnnASP02_STRING;
	rsPhoneType.Source = "{call dbo.cp_phone_type2("+ Request.Form("MM_recordId") + ",'" + Abbreviation + "','" + Description + "',0,'E',0)}";
	rsPhoneType.CursorType = 0;
	rsPhoneType.CursorLocation = 2;
	rsPhoneType.LockType = 3;
	rsPhoneType.Open();
	Response.Redirect("m018q0313.asp");
}

var rsPhoneType = Server.CreateObject("ADODB.Recordset");
rsPhoneType.ActiveConnection = MM_cnnASP02_STRING;
rsPhoneType.Source = "{call dbo.cp_phone_type2("+ Request.QueryString("intPhone_type_id") + ",'','',1,'Q',0)}";
rsPhoneType.CursorType = 0;
rsPhoneType.CursorLocation = 2;
rsPhoneType.LockType = 3;
rsPhoneType.Open();
%>
<html>
<head>
	<title>Update Phone Type Lookup</title>
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
				document.frm0313.reset();
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
		if (Trim(document.frm0313.Description.value)==""){
			alert("Enter Description.");
			document.frm0313.Description.focus();
			return ;		
		}
		document.frm0313.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0313.Description.focus();">
<form name="frm0313" method="POST" action="<%=MM_editAction%>">
<h5>Update Phone Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="<%=(rsPhoneType.Fields.Item("chvName").Value)%>" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Abbreviation:</td>
		<td nowrap><input type="text" name="Abbreviation" value="<%=(rsPhoneType.Fields.Item("chvAbbrev").Value)%>" tabindex="2" size="5" maxlength="5" accesskey="L"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="3" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="5" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=rsPhoneType.Fields.Item("intPhone_type_id").Value %>">
</form>
</body>
</html>
<%
rsPhoneType.Close();
%>