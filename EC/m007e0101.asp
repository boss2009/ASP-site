<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_editAction += "?" + Request.QueryString;
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

if (Request.Form("MM_action") == "insert"){
	var AbstractClassName = String(Request.Form("AbstractClassName")).replace(/'/g, "''");		
	var rsAbstractClass = Server.CreateObject("ADODB.Recordset");
	rsAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
	rsAbstractClass.Source = "{call dbo.cp_Update_Eqp_Class(" + Request.Form("ClassID") + ",'" + AbstractClassName + "',0,1," + Session("insStaff_id") + ",'0',0,0,0,1,0,0,'A',0)}";
	rsAbstractClass.CursorType = 0;
	rsAbstractClass.CursorLocation = 2;
	rsAbstractClass.LockType = 3;
	rsAbstractClass.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Update");
}

var rsAbstractClass = Server.CreateObject("ADODB.Recordset");
rsAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW(" + Request.QueryString("ClassID") + ",'A',1)}";	
rsAbstractClass.CursorType = 0;
rsAbstractClass.CursorLocation = 2;
rsAbstractClass.LockType = 3;
rsAbstractClass.Open();
%>
<html>
<head>
	<title>Abstract Class</title>
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
	
	function DeleteClass(){
		if (confirm("Delete This Class?")) {
			document.frm0101.MM_action.value="delete";
			document.frm0101.submit();
		} 		
	}
	</script>
</head>
<body onLoad="document.frm0101.AbstractClassName.focus();"> 
<form action="<%=MM_editAction%>" method="POST" name="frm0101">
<h5>Abstract Class</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Abstract Class Name:</td>
		<td nowrap><input type="text" name="AbstractClassName" maxlength="50" value="<%=(rsAbstractClass.Fields.Item("chvName").Value)%>" size="50" tabindex="1" accesskey="F" ></td>
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
<input type="hidden" name="ClassID" value="<%=Request.QueryString("classid")%>">
<input type="hidden" name="MM_action" value="update">
</form>
</body>
</html>
