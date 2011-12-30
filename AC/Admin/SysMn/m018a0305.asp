<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_Insert")) == "true") {	
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var Abbreviation = String(Request.Form("Abbreviation")).replace(/'/g, "''");					
	var rsProvince = Server.CreateObject("ADODB.Recordset");
	rsProvince.ActiveConnection = MM_cnnASP02_STRING;
	rsProvince.Source = "{call dbo.cp_prov_state2(0,'"+Abbreviation+"','"+ Description + "',"+Request.Form("Country")+",0,'A',0)}";
	rsProvince.CursorType = 0;
	rsProvince.CursorLocation = 2;
	rsProvince.LockType = 3;
	rsProvince.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Province/State Lookup</title>
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
				document.frm0305.reset();
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
		if (Trim(document.frm0305.Description.value)==""){
			alert("Enter Description.");
			document.frm0305.Description.focus();
			return ;		
		}
		document.frm0305.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0305.Description.focus();">
<form name="frm0305" method="POST" action="<%=MM_editAction%>">
<h5>New Province/State Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="50" size="30" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td>Abbreviation:</td>
		<td><input type="text" name="Abbreviation" tabindex="2" size="2" maxlength="2"></td>
	</tr>
    <tr> 
		<td>Country:</td>
		<td><select name="Country" tabindex="3" accesskey="L">
			<option value="1">Canada
			<option value="2">United States
		</select></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="3" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="4" onClick="window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>