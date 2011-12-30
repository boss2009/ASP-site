<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var rsCourseType = Server.CreateObject("ADODB.Recordset");
	rsCourseType.ActiveConnection = MM_cnnASP02_STRING;
	rsCourseType.Source = "{call dbo.cp_course_type(0,'"+ Description + "',0,'A',0)}";
	rsCourseType.CursorType = 0;
	rsCourseType.CursorLocation = 2;
	rsCourseType.LockType = 3;
	rsCourseType.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Course Type</title>
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
		if (Trim(document.frm0326.Description.value)=="") {
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
<h5>New Course Type</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
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