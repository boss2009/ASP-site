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
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var rsWorkType = Server.CreateObject("ADODB.Recordset");
	rsWorkType.ActiveConnection = MM_cnnASP02_STRING;
	rsWorkType.Source = "{call dbo.cp_Work_Type("+ Request.Form("MM_recordId") + ",'" + Request.Form("Description") + "'," + IsActive + ",0,'E',0)}";
	rsWorkType.CursorType = 0;
	rsWorkType.CursorLocation = 2;
	rsWorkType.LockType = 3;
	rsWorkType.Open();
	Response.Redirect("m018q03125.asp");
}

var rsWorkType = Server.CreateObject("ADODB.Recordset");
rsWorkType.ActiveConnection = MM_cnnASP02_STRING;
rsWorkType.Source = "{call dbo.cp_Work_Type("+ Request.QueryString("intWork_type_id") + ",'',0,1,'Q',0)}";
rsWorkType.CursorType = 0;
rsWorkType.CursorLocation = 2;
rsWorkType.LockType = 3;
rsWorkType.Open();
%>
<html>
<head>
	<title>Update Work Type Lookup</title>
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
				document.frm03125.reset();
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
		if (Trim(document.frm03125.Description.value)==""){
			alert("Enter Work Type Name.");
			document.frm03125.Description.focus();
			return ;		
		}
		document.frm03125.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03125.Description.focus();">
<form name="frm03125" method="POST" action="<%=MM_editAction%>">
<h5>Update Work Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="<%=(rsWorkType.Fields.Item("chvWork_type_desc").Value)%>" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Is Active:</td>
		<td nowrap><input type="checkbox" name="IsActive" <%=((rsWorkType.Fields.Item("bitis_active").Value == 1)?"CHECKED":"")%> value="1" tabindex="2" accesskey="L" class="chkstyle"></td>
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
<input type="hidden" name="MM_recordId" value="<%= rsWorkType.Fields.Item("intWork_type_id").Value %>">
</form>
</body>
</html>
<%
rsWorkType.Close();
%>