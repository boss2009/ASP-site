<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}
if (String(Request("MM_update")) != "undefined" && String(Request("MM_recordId")) != "undefined") {	
	var AreaCode = String(Request.Form("AreaCode")).replace(/'/g, "''");	
	var IsLocal = ((Request.Form("IsLocal")=="1") ? "1":"0");
	var rsAreaCode = Server.CreateObject("ADODB.Recordset");
	rsAreaCode.ActiveConnection = MM_cnnASP02_STRING;
	rsAreaCode.Source = "{call dbo.cp_area_code("+ Request.Form("MM_recordId") + ",'" + Request.Form("AreaCode") + "'," + IsLocal + ",0,'E',0)}";
	rsAreaCode.CursorType = 0;
	rsAreaCode.CursorLocation = 2;
	rsAreaCode.LockType = 3;
	rsAreaCode.Open();
	Response.Redirect("m018q03129.asp");
}

var rsAreaCode = Server.CreateObject("ADODB.Recordset");
rsAreaCode.ActiveConnection = MM_cnnASP02_STRING;
rsAreaCode.Source = "{call dbo.cp_area_code("+ Request.QueryString("intAC_Id") + ",'',0,1,'Q',0)}";
rsAreaCode.CursorType = 0;
rsAreaCode.CursorLocation = 2;
rsAreaCode.LockType = 3;
rsAreaCode.Open();
%>
<html>
<head>
	<title>Update Area Code Lookup</title>
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
				document.frm03129.reset();
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
		if (Trim(document.frm03129.AreaCode.value)==""){
			alert("Enter Area Code.");
			document.frm03129.AreaCode.focus();
			return ;		
		}
		document.frm03129.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03129.AreaCode.focus();">
<form name="frm03129" method="POST" action="<%=MM_editAction%>">
<h5>Update Area Code Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Area Code:</td>
		<td nowrap><input type="text" name="AreaCode" value="<%=(rsAreaCode.Fields.Item("chvAC_num").Value)%>" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Is Local:</td>
		<td nowrap><input type="checkbox" name="IsLocal" <%=((rsAreaCode.Fields.Item("bitIs_Local").Value == 1)?"CHECKED":"")%> value="1" tabindex="2" accesskey="L" class="chkstyle"></td>
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
<input type="hidden" name="MM_recordId" value="<%= rsAreaCode.Fields.Item("intAC_Id").Value %>">
</form>
</body>
</html>
<%
rsAreaCode.Close();
%>