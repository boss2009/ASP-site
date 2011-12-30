<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_actionAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_actionAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_action"))=="delete"){
	var rsDeleteClass = Server.CreateObject("ADODB.Recordset");
	rsDeleteClass.ActiveConnection = MM_cnnASP02_STRING;
	rsDeleteClass.Source = "{call dbo.cp_Delete_Eqp_Class(" + Request.QueryString("ClassID") + ",0)}";	
	rsDeleteClass.CursorType = 0;
	rsDeleteClass.CursorLocation = 2;
	rsDeleteClass.LockType = 3;
	rsDeleteClass.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Deleted");	
}

if (String(Request.Form("MM_action"))=="update"){
	var SubAbstractClassName = String(Request.Form("SubAbstractClassName")).replace(/'/g, "''");		
	var rsSubAbstractClass = Server.CreateObject("ADODB.Recordset");
	rsSubAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
	rsSubAbstractClass.Source = "{call dbo.cp_Update_Eqp_Class(" + Request.Form("ClassID") + ",'" + SubAbstractClassName + "'," + Request.Form("ParentClass") + ",1," + Session("insStaff_id") + ",'0',0,0,0,1,0,0,'S',0)}";
	rsSubAbstractClass.CursorType = 0;
	rsSubAbstractClass.CursorLocation = 2;
	rsSubAbstractClass.LockType = 3;
	rsSubAbstractClass.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Update");
}

var rsSubAbstractClass = Server.CreateObject("ADODB.Recordset");
rsSubAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsSubAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW(" + Request.QueryString("ClassID") + ",'S',1)}";	
rsSubAbstractClass.CursorType = 0;
rsSubAbstractClass.CursorLocation = 2;
rsSubAbstractClass.LockType = 3;
rsSubAbstractClass.Open();
%>
<html>
<head>
	<title>Sub Abstract Class</title>
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
		if (Trim(document.frm0102.SubAbstractClassName.value)==""){
			alert("Enter Sub Abstract Class Name.");
			document.frm0102.SubAbstractClassName.focus();
			return ;
		}
		document.frm0102.submit();
	}

	function DeleteClass(){
		if (confirm("Delete This Class?")) {
			document.frm0102.MM_action.value="delete";
			document.frm0102.submit();
		} 		
	}	
	</script>
</head>
<body onLoad="document.frm0102.ParentClass.focus();"> 
<form action="<%=MM_actionAction%>" method="POST" name="frm0102">
<h5>Sub Abstract Class</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Parent Class Name:</td>
		<td nowrap><select name="ParentClass" tabindex="1" accesskey="F">
			<option value=<%=(rsSubAbstractClass.Fields.Item("insAbsCls_id").Value)%>><%=rsSubAbstractClass.Fields.Item("chvAbsClsName").Value%>
		</select></td>
	</tr>
	<tr>
		<td nowrap>Sub Abstract Class Name:</td>
		<td nowrap><input type="text" name="SubAbstractClassName" maxlength="50" value="<%=(rsSubAbstractClass.Fields.Item("chvSubAbsClsName").Value)%>" size="50" tabindex="2" accesskey="L"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="2" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="3" class="btnstyle"></td>
<% 
	if (Session("MM_UserAuthorization") >= 5 ){
%>		
		<td><input type="button" value="Delete" onClick="DeleteClass();" tabindex="4" class="btnstyle"></td>		
<% 
	} 
%>
	</tr>
</table>
<input type="hidden" name="ClassID" value="<%=Request.QueryString("ClassID")%>">
<input type="hidden" name="MM_action" value="update">
</form>
</body>
</html>