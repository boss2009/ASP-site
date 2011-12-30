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
	var rsCourseType = Server.CreateObject("ADODB.Recordset");
	rsCourseType.ActiveConnection = MM_cnnASP02_STRING;
	rsCourseType.Source = "{call dbo.cp_course_type("+ Request.Form("MM_recordId") + ",'" + Description + "',0,'E',0)}";
	rsCourseType.CursorType = 0;
	rsCourseType.CursorLocation = 2;
	rsCourseType.LockType = 3;
	rsCourseType.Open();
	Response.Redirect("m018q0326.asp");
}

var rsCourseType = Server.CreateObject("ADODB.Recordset");
rsCourseType.ActiveConnection = MM_cnnASP02_STRING;
rsCourseType.Source = "{call dbo.cp_course_type("+ Request.QueryString("insCourse_id") + ",'',1,'Q',0)}";
rsCourseType.CursorType = 0;
rsCourseType.CursorLocation = 2;
rsCourseType.LockType = 3;
rsCourseType.Open();
%>
<html>
<head>
	<title>Update Course Type Lookup</title>
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
				document.frm0326.reset();
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
		if (Trim(document.frm0326.Description.value)==""){
			alert("Enter Description.");
			document.frm0326.Description.focus();
			return ;		
		}
		document.frm0326.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0326.Description.focus();">
<form name="frm0326" method="POST" action="<%=MM_editAction%>">
<h5>Update Course Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsCourseType.Fields.Item("chvcourse_name").Value)%>" maxlength="50" size="30" tabindex="1" accesskey="F" ></td>
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
<input type="hidden" name="MM_recordId" value="<%=rsCourseType.Fields.Item("insCourse_id").Value %>">
</form>
</body>
</html>
<%
rsCourseType.Close();
%>