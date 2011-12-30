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
	var SubAbstractClassName = String(Request.Form("SubAbstractClassName")).replace(/'/g, "''");		
	var rsSubAbstractClass = Server.CreateObject("ADODB.Recordset");
	rsSubAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
	rsSubAbstractClass.Source = "{call dbo.cp_Insert_Eqp_Class('"+SubAbstractClassName+ "',1," + Request.Form("ParentClass") + ",1,"+ Session("insStaff_id") + ",'" + CurrentDate() + "',0,0,'S',0)}";
	rsSubAbstractClass.CursorType = 0;
	rsSubAbstractClass.CursorLocation = 2;
	rsSubAbstractClass.LockType = 3;
	rsSubAbstractClass.Open();
	Response.Redirect("InsertSuccessful.html");
}

var rsAbstractClass = Server.CreateObject("ADODB.Recordset");
rsAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW(0,'A',0)}";
rsAbstractClass.CursorType = 0;
rsAbstractClass.CursorLocation = 2;
rsAbstractClass.LockType = 3;
rsAbstractClass.Open();
%>
<html>
<head>
	<title>New Sub Abstract Class</title>
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
	</script>
</head>
<body onLoad="document.frm0102.SubAbstractClassName.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0102">
<h5>New Sub Abstract Class</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Add Under Abstract Class:</td>
		<td nowrap><select name="ParentClass" accesskey="F" tabindex="1">
			<% 
			while (!rsAbstractClass.EOF){
			%>
				<option value="<%=rsAbstractClass.Fields.Item("insEquip_Class_id").Value%>" <%=((rsAbstractClass.Fields.Item("insEquip_Class_id").Value==Request.QueryString("ClassID"))?"SELECTED":"")%>><%=rsAbstractClass.Fields.Item("chvName").Value%>
			<% 
				rsAbstractClass.MoveNext(); 
			}
			%>
		</select></td>
	</tr>
    <tr> 
		<td nowrap>Sub Abstract Class Name:</td>
		<td nowrap><input type="text" name="SubAbstractClassName" maxlength="50" size="50" tabindex="2" accesskey="L" ></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="4" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsAbstractClass.Close();
%>