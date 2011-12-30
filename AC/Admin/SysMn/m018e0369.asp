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
	var rsDuration = Server.CreateObject("ADODB.Recordset");
	rsDuration.ActiveConnection = MM_cnnASP02_STRING;
	rsDuration.Source = "{call dbo.cp_duratn_type2("+ Request.Form("MM_recordId") + ",'" + Description + "','" + Abbreviation + "',0,'E',0)}";
	rsDuration.CursorType = 0;
	rsDuration.CursorLocation = 2;
	rsDuration.LockType = 3;
	rsDuration.Open();
	Response.Redirect("m018q0369.asp");
}

var rsDuration = Server.CreateObject("ADODB.Recordset");
rsDuration.ActiveConnection = MM_cnnASP02_STRING;
rsDuration.Source = "{call dbo.cp_duratn_type2("+ Request.QueryString("insDuratn_type_id") + ",'','',1,'Q',0)}";
rsDuration.CursorType = 0;
rsDuration.CursorLocation = 2;
rsDuration.LockType = 3;
rsDuration.Open();
%>
<html>
<head>
	<title>Update Duration Type Lookup</title>
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
<h5>Update Duration Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsDuration.Fields.Item("chvDuratn_desc").Value)%>" maxlength="50" size="30" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td>Abbreviation:</td>
		<td><input type="text" name="Abbreviation" value="<%=(rsDuration.Fields.Item("chrAbbrev").Value)%>" tabindex="2" size="10" maxlength="10" accesskey="L"></td>
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
<input type="hidden" name="MM_recordId" value="<%=rsDuration.Fields.Item("insDuratn_type_id").Value %>">
</form>
</body>
</html>
<%
rsDuration.Close();
%>