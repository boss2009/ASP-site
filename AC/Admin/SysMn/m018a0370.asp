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
	var IsTax = ((Request.Form("IsTax")=="on") ? "1":"0");
	var rsChargeRate = Server.CreateObject("ADODB.Recordset");
	rsChargeRate.ActiveConnection = MM_cnnASP02_STRING;
	rsChargeRate.Source = "{call dbo.cp_charge_rate(0,'"+ Description + "'," + IsTax + "," + Request.Form("Percentage") + "0,'A',0)}";
	rsChargeRate.CursorType = 0;
	rsChargeRate.CursorLocation = 2;
	rsChargeRate.LockType = 3;
	rsChargeRate.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Charge Rate</title>
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
		if (Trim(document.frm0370.Description.value)==""){
			alert("Enter Description.");
			document.frm0370.Description.focus();
			return ;		
		}
		if (isNaN(document.frm0370.Percentage.value)){
			alert("Invalid Percentage.");
			document.frm0370.Percentage.focus();
			return ;
		}
		document.frm0370.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0370.Description.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0370">
<h5>New Charge Rate</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td nowrap>Is Tax:</td> 
		<td nowrap><input type="checkbox" name="IsTax" value="1" tabindex="2" class="chkstyle"></td>
	</tr>
	<tr>
		<td nowrap>Percentage:</td>		
		<td nowrap><input type="text" name="Percentage" maxlength="4" size="4" tabindex="3" onKeypress="AllowNumericOnly();" accesskey="L" style="text-align: right">%</td>
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