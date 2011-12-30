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
	var IsHardware = ((Request.Form("IsHardware")=="1") ? "1":"0");
	var IsDisabilityDocumentation = ((Request.Form("IsDisabilityDocumentation")=="1") ? "1":"0");	
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var rsHWLocation = Server.CreateObject("ADODB.Recordset");
	rsHWLocation.ActiveConnection = MM_cnnASP02_STRING;
	rsHWLocation.Source = "{call dbo.cp_hw_location(0,'"+Description+"',"+IsHardware+","+IsDisabilityDocumentation+",0,'A',0)}";
	rsHWLocation.CursorType = 0;
	rsHWLocation.CursorLocation = 2;
	rsHWLocation.LockType = 3;
	rsHWLocation.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Hardware Location Lookup</title>
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
		if (Trim(document.frm03108.Description.value)=="") {
			alert("Enter Description.");
			document.frm03108.Description.focus();
			return ;
		}
		document.frm03108.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03108.Description.focus();">
<form name="frm03108" method="POST" action="<%=MM_editAction%>">
<h5>New Hardware Location Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" maxlength="40" size="20" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td nowrap>Is Hardware:</td>
		<td nowrap><input type="checkbox" name="IsHardware" value="1" tabindex="2" class="chkstyle"></td>
    </tr>
    <tr> 
		<td nowrap>Is Disability Documentation:</td>
		<td nowrap><input type="checkbox" name="IsDisabilityDocumentation" value="1" tabindex="3" accesskey="L" class="chkstyle"></td>
    </tr>		
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="5" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>