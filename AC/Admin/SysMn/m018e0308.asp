<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var rsStatus = Server.CreateObject("ADODB.Recordset");
	rsStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsStatus.Source = "{call dbo.cp_AC_StdStatus2(" + Request.QueryString("insStdnt_status_id") + ",0,'" + Description + "'," + IsActive + ",1,0,'E',0)}";
	rsStatus.CursorType = 0;
	rsStatus.CursorLocation = 2;
	rsStatus.LockType = 3;
	rsStatus.Open();
	Response.Redirect("m018q0308.asp");
}

var rsStatus = Server.CreateObject("ADODB.Recordset");
rsStatus.ActiveConnection = MM_cnnASP02_STRING;
rsStatus.Source = "{call dbo.cp_AC_StdStatus2(" + Request.QueryString("insStdnt_status_id") + ",0,'',0,0,1,'Q',0)}";
rsStatus.CursorType = 0;
rsStatus.CursorLocation = 2;
rsStatus.LockType = 3;
rsStatus.Open();
%>
<html>
<head>
	<title>Update Status Lookup</title>
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
				document.frm0308.reset();
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
		if (Trim(document.frm0308.Description.value)==""){
			alert("Enter Description.");
			document.frm0308.Description.focus();
			return ;
		}
		document.frm0308.submit();
	}
	</script>
</head>
<body onLoad="document.frm0308.Description.focus();">
<form name="frm0308" method="POST" action="<%=MM_editAction%>">
<h5>Update Status Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="<%=(rsStatus.Fields.Item("chvStdnt_Status").Value)%>" maxlength="40" size="20" tabindex="1" accesskey="F"></td>
    </tr>
    <tr>
		<td nowrap>Is Active:</td>
		<td nowrap><input type="checkbox" name="IsActive" <%=((rsStatus.Fields.Item("bitIs_adult_status").Value == 1)?"CHECKED":"")%> value="1" tabindex="2" accesskey="L" class="chkstyle"></td>
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
<input type="hidden" name="MM_recordId" value="<%= rsStatus.Fields.Item("insStdnt_status_id").Value %>">
</form>
</body>
</html>
<%
rsStatus.Close();
%>