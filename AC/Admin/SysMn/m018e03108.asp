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
	var IsHardware = ((Request.Form("IsHardware")=="1") ? "1":"0");
	var IsDisabilityDocumentation = ((Request.Form("IsDisabilityDocumentation")=="1") ? "1":"0");	
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var rsHardwareLocation = Server.CreateObject("ADODB.Recordset");
	rsHardwareLocation.ActiveConnection = MM_cnnASP02_STRING;
	rsHardwareLocation.Source = "{call dbo.cp_hw_location("+Request.QueryString("insLocation_id")+",'"+Description+"',"+IsHardware+","+IsDisabilityDocumentation+",0,'E',0)}";
	rsHardwareLocation.CursorType = 0;
	rsHardwareLocation.CursorLocation = 2;
	rsHardwareLocation.LockType = 3;
	rsHardwareLocation.Open();
	Response.Redirect("m018q03108.asp");
}

var rsHardwareLocation = Server.CreateObject("ADODB.Recordset");
rsHardwareLocation.ActiveConnection = MM_cnnASP02_STRING;
rsHardwareLocation.Source = "{call dbo.cp_hw_location("+Request.QueryString("insLocation_id")+",'',0,0,1,'Q',0)}";
rsHardwareLocation.CursorType = 0;
rsHardwareLocation.CursorLocation = 2;
rsHardwareLocation.LockType = 3;
rsHardwareLocation.Open();
%>
<html>
<head>
	<title>Update Hardware Location Lookup</title>
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
				document.frm03108.reset();
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
		if (Trim(document.frm03108.Description.value)==""){
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
<h5>Update Hardware Location Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsHardwareLocation.Fields.Item("chvLocation_Desc").Value)%>" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td>Is Hardware:</td>
		<td><input type="checkbox" name="IsHardware" value="1" <%=((rsHardwareLocation.Fields.Item("bitIs_HW").Value=="1")?"CHECKED":"")%> tabindex="2" class="chkstyle"></td>
    </tr>
    <tr> 
		<td>Is Disability Documentation:</td>
		<td><input type="checkbox" name="IsDisabilityDocumentation" value="1" <%=((rsHardwareLocation.Fields.Item("bitIs_Dsbty_Doc").Value=="1")?"CHECKED":"")%> tabindex="3" accesskey="L" class="chkstyle"></td>
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
<input type="hidden" name="MM_recordId" value="<%= rsHardwareLocation.Fields.Item("insLocation_id").Value %>">
</form>
</body>
</html>
<%
rsHardwareLocation.Close();
%>