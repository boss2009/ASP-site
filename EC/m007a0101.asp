<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (Request.Form("MM_insert") == "true"){
	var AbstractClassName = String(Request.Form("AbstractClassName")).replace(/'/g, "''");	
	var rsAbstractClass = Server.CreateObject("ADODB.Recordset");
	rsAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
	rsAbstractClass.Source = "{call dbo.cp_Insert_Eqp_Class('"+AbstractClassName+"',1,0,1,"+ Session("insStaff_id") + ",'" + CurrentDate() + "',0,0,'A',0)}";
	rsAbstractClass.CursorType = 0;
	rsAbstractClass.CursorLocation = 2;
	rsAbstractClass.LockType = 3;
	rsAbstractClass.Open();
	Response.Redirect("InsertSuccessful.html");
}
%>
<html>
<head>
	<title>New Abstract Class</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
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
		if (Trim(document.frm0101.AbstractClassName.value)==""){
			alert("Enter Abstract Class Name.");
			document.frm0101.AbstractClassName.focus();
			return ;
		}
		document.frm0101.submit();
	}
	</script>
</head>
<body onLoad="document.frm0101.AbstractClassName.focus();">
<form action="<%=MM_editAction%>" method="POST" name="frm0101">
<h5>New Abstract Class</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Abstract Class Name:</td>
		<td nowrap><input type="text" name="AbstractClassName" maxlength="50" size="50" tabindex="1" accesskey="F"></td>
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