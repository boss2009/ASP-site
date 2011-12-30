<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (Request.Form("MM_insert") == "true"){
	var ConcreteClassName = String(Request.Form("ConcreteClassName")).replace(/'/g, "''");		
	var ConcreteClass = Server.CreateObject("ADODB.Recordset");
	ConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
	ConcreteClass.Source = "{call dbo.cp_Insert_Eqp_Class('" + ConcreteClassName + "',0," + Request.Form("ParentID") + "," + Request.Form("ClassStatus") +"," + Session("insStaff_id") + ",'" + CurrentDateTime() + "','" + Request.Form("ModelNumber") + "','" + Request.Form("SubjectTo") +"','C',0)}";
	ConcreteClass.CursorType = 0;
	ConcreteClass.CursorLocation = 2;
	ConcreteClass.LockType = 3;
	ConcreteClass.Open();
	Response.Redirect("InsertSuccessful.html");
}

var rsSubAbstractClass = Server.CreateObject("ADODB.Recordset");
rsSubAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsSubAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW(" + Request.QueryString("ParentID") + ",'S',1)}";
rsSubAbstractClass.CursorType = 0;
rsSubAbstractClass.CursorLocation = 2;
rsSubAbstractClass.LockType = 3;
rsSubAbstractClass.Open();
%>
<html>
<head>
	<title>New Concrete Class</title>
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
		if (Trim(document.frm0103.ConcreteClassName.value)==""){
			alert("Enter Concrete Class Name.");
			document.frm0103.ConcreteClassName.focus();
			return ;
		}
		document.frm0103.submit();
	}
	</script>
</head>
<body onLoad="document.frm0103.ConcreteClassName.focus();"> 
<form action="<%=MM_editAction%>" method="POST" name="frm0103">
<h5>New Concrete Class</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Abstract Class Name:</td>
		<td nowrap><input type="text" name="AbstractClassName" maxlength="50" value="<%=rsSubAbstractClass.Fields.Item("chvAbsClsName").Value%>" size="50" readonly tabindex="1" accesskey="F" ></td>
	</tr>
	<tr>
		<td nowrap>Sub Abstract Class Name:</td>
		<td nowrap><input type="text" name="SubAbstractClassName" maxlength="50" value="<%=rsSubAbstractClass.Fields.Item("chvSubAbsClsName").Value%>" size="50" readonly tabindex="2" ></td>
	<tr>
		<td nowrap>Concrete Class Name:</td>
		<td nowrap><input type="text" name="ConcreteClassName" maxlength="50" value="" size="50" tabindex="3" ></td>
	</tr>
	<tr>
		<td nowrap>Model Number:</td>
		<td nowrap><input type="text" name="ModelNumber" value="" maxlength="50" size="15" tabindex="4" ></td>
	</tr>	
	<tr>
		<td nowrap>Subject To:</td>
		<td nowrap><select name="SubjectTo" tabindex="5">
			<option value="0">No Tax
			<option value="1">PST
			<option value="2">GST
			<option value="3">PST/GST
		</select></td>
	</tr>
	<tr>
		<td nowrap>Class Status:</td>
		<td nowrap><select name="ClassStatus" tabindex="6" accesskey="L">
			<option value="1">Active
			<option value="0">Inactive
		</select></td>
	</tr>	
<!--
	<tr>
		<td nowrap valign="top">Notes:</td>
		<td><textarea name="Notes" tabindex="7" rows="4" cols="65" accesskey="L"></textarea></td>
	</tr>
-->
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="8" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="top.window.close();" tabindex="9" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="SubAbstractClassID" value="<%=rsSubAbstractClass.Fields.Item("insSubAbsCls_id").Value%>">
<input type="hidden" name="ParentID" value="<%=Request.QueryString("ParentID")%>">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsSubAbstractClass.Close();
%>