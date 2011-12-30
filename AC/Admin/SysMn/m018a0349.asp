<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
// set the form action variable
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var Description = String(Request.Form("Description")).replace(/'/g, "''");	
	var ForClientModule = ((Request.Form("ForClientModule")=="1") ? "1":"0");	
	var ForServiceRequest = ((Request.Form("ForServiceRequest")=="1") ? "1":"0");	
	var ForInstitutionModule = ((Request.Form("ForInstitutionModule")=="1") ? "1":"0");	
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");		
	var rsServiceType = Server.CreateObject("ADODB.Recordset");
	rsServiceType.ActiveConnection = MM_cnnASP02_STRING;
	rsServiceType.Source = "{call dbo.cp_service_type2(0,'"+Description+"',"+ForClientModule+","+ForServiceRequest+","+ForInstitutionModule+","+IsActive+",0,'A',0)}";
	rsServiceType.CursorType = 0;
	rsServiceType.CursorLocation = 2;
	rsServiceType.LockType = 3;
	rsServiceType.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Service Type</title>
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
		if (Trim(document.frm0349.Description.value)==""){
			alert("Enter Description.");
			document.frm0349.Description.focus();
			return ;		
		}
		document.frm0349.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0349.Description.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0349">
<h5>New Service Type</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="40" size="40" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td>For Client Module:</td>
		<td><input type="checkbox" name="ForClientModule" value="1" tabindex="2" class="chkstyle"></td>
    </tr>
    <tr> 
		<td>For Institution Module:</td>
		<td><input type="checkbox" name="ForInstitutionModule" value="1" tabindex="3" class="chkstyle"></td>        
    </tr>
	<tr>			  
		<td>For Service Request:</td>
		<td><input type="checkbox" name="ForServiceRequest" value="1" tabindex="4" class="chkstyle"></td>
    </tr>
    <tr> 
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" value="1" tabindex="5" accesskey="L" class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="6" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="7" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>