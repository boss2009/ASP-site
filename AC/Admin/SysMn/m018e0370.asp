<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true"){
	var Description = String(Request.Form("Description")).replace(/'/g, "''");	
	var IsTax = ((Request.Form("IsTax")=="1") ? "1":"0");
	var rsChargeRate = Server.CreateObject("ADODB.Recordset");
	rsChargeRate.ActiveConnection = MM_cnnASP02_STRING;
	rsChargeRate.Source = "{call dbo.cp_charge_rate("+ Request.Form("MM_recordId") + ",'" + Request.Form("Description") + "'," + IsTax + ","+Request.Form("Percentage")+",0,'E',0)}";
	rsChargeRate.CursorType = 0;
	rsChargeRate.CursorLocation = 2;
	rsChargeRate.LockType = 3;
	rsChargeRate.Open();
	Response.Redirect("m018q0370.asp");
}

var rsChargeRate = Server.CreateObject("ADODB.Recordset");
rsChargeRate.ActiveConnection = MM_cnnASP02_STRING;
rsChargeRate.Source = "{call dbo.cp_charge_rate("+ Request.QueryString("intCharge_Rate_id") + ",'',0,.0,1,'Q',0)}";
rsChargeRate.CursorType = 0;
rsChargeRate.CursorLocation = 2;
rsChargeRate.LockType = 3;
rsChargeRate.Open();
%>
<html>
<head>
	<title>Update Charge Rate Lookup</title>
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
			document.frm0370.reset();
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
		if (Trim(document.frm0370.Description.value)==""){
			alert("Enter Description.");
			document.frm0370.Description.focus();
			return ;		
		}
		if (isNaN(document.frm0370.Percentage.value)){
			alert("Invalid Percentage.");
			document.frm0370.Percentage.focus();
		}
		document.frm0370.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0370.Description.focus();">
<form name="frm0370" method="POST" action="<%=MM_editAction%>">
<h5>Update Charge Rate Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsChargeRate.Fields.Item("chvCharge_Item_Desc").Value)%>" maxlength="50" size="30" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td>Is Tax:</td>
		<td><input type="checkbox" name="IsTax" <%=((rsChargeRate.Fields.Item("bitIsTax").Value == 1)?"CHECKED":"")%> value="1" tabindex="2" class="chkstyle"></td>
	</tr>
    <tr> 
		<td>Percentage:</td>
		<td><input type="text" name="Percentage" value="<%=(rsChargeRate.Fields.Item("fltPercentage").Value)%>" maxlength="4" size="4" tabindex="3" style="text-align: right" accesskey="L" onKeypress="AllowNumericOnly();">%</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="4" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="5" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="6" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsChargeRate.Fields.Item("intCharge_Rate_id").Value %>">
</form>
</body>
</html>
<%
rsChargeRate.Close();
%>