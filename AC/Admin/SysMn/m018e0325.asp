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
	var rsSemester = Server.CreateObject("ADODB.Recordset");
	rsSemester.ActiveConnection = MM_cnnASP02_STRING;
	rsSemester.Source = "{call dbo.cp_semester2("+ Request.Form("MM_recordId") + ",'" + Description + "',0,'E',0)}";
	rsSemester.CursorType = 0;
	rsSemester.CursorLocation = 2;
	rsSemester.LockType = 3;
	rsSemester.Open();
	Response.Redirect("m018q0325.asp");
}

var rsSemester = Server.CreateObject("ADODB.Recordset");
rsSemester.ActiveConnection = MM_cnnASP02_STRING;
rsSemester.Source = "{call dbo.cp_semester2("+ Request.QueryString("insSmstr_id") + ",'',1,'Q',0)}";
rsSemester.CursorType = 0;
rsSemester.CursorLocation = 2;
rsSemester.LockType = 3;
rsSemester.Open();
%>
<html>
<head>
	<title>Update Semester Lookup</title>
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
				document.frm0325.reset();
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
		if (Trim(document.frm0325.Description.value)==""){
			alert("Enter Description.");
			document.frm0325.Description.focus();
			return ;		
		}
		document.frm0325.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0325.Description.focus();">
<form name="frm0325" method="POST" action="<%=MM_editAction%>">
<h5>Update Semester Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsSemester.Fields.Item("chvsmstr_name").Value)%>" maxlength="50" size="30" tabindex="1" accesskey="F"></td>
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
<input type="hidden" name="MM_recordId" value="<%=rsSemester.Fields.Item("insSmstr_id").Value %>">
</form>
</body>
</html>
<%
rsSemester.Close();
%>