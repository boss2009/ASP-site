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
	var ForClientModule = ((Request.Form("ForClientModule")=="1") ? "1":"0");	
	var ForServiceRequest = ((Request.Form("ForServiceRequest")=="1") ? "1":"0");	
	var ForInstitutionModule = ((Request.Form("ForInstitutionModule")=="1") ? "1":"0");	
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");		
	var rsServiceType = Server.CreateObject("ADODB.Recordset");
	rsServiceType.ActiveConnection = MM_cnnASP02_STRING;
	rsServiceType.Source = "{call dbo.cp_service_type2("+Request.QueryString("insService_type_id")+",'"+Description+"',"+ForClientModule+","+ForServiceRequest+","+ForInstitutionModule+","+IsActive+",0,'E',0)}";
	rsServiceType.CursorType = 0;
	rsServiceType.CursorLocation = 2;
	rsServiceType.LockType = 3;
	rsServiceType.Open();
	Response.Redirect("m018q0349.asp");
}

var rsServiceType = Server.CreateObject("ADODB.Recordset");
rsServiceType.ActiveConnection = MM_cnnASP02_STRING;
rsServiceType.Source = "{call dbo.cp_service_type2("+Request.QueryString("insService_type_id")+",'',0,0,0,0,1,'Q',0)}";
rsServiceType.CursorType = 0;
rsServiceType.CursorLocation = 2;
rsServiceType.LockType = 3;
rsServiceType.Open();
%>
<html>
<head>
	<title>Update Service Type Lookup</title>
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
				document.frm0349.reset();
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
<form name="frm0349" method="POST" action="<%=MM_editAction%>">
<h5>Update Service Type Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" value="<%=(rsServiceType.Fields.Item("chvService_type").Value)%>" maxlength="50" size="30" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td>For Client Module:</td>
		<td><input type="checkbox" name="ForClientModule" <%=((rsServiceType.Fields.Item("bitis_adult").Value == 1)?"CHECKED":"")%> value="1" tabindex="2" class="chkstyle"></td>
    </tr>
    <tr> 
		<td>For Institution Module:</td>
		<td><input type="checkbox" name="ForInstitutionModule" <%=((rsServiceType.Fields.Item("bitis_School_Class").Value == 1)?"CHECKED":"")%> value="1" tabindex="3" class="chkstyle"></td>        
    </tr>
	<tr>			  
		<td>For Service Request:</td>
		<td><input type="checkbox" name="ForServiceRequest" <%=((rsServiceType.Fields.Item("bitis_req_service").Value == 1)?"CHECKED":"")%> value="1" tabindex="4" class="chkstyle"></td>
    </tr>
    <tr> 
		<td>Is Active:</td>
		<td><input type="checkbox" name="IsActive" <%=((rsServiceType.Fields.Item("bitis_active").Value == 1)?"CHECKED":"")%> value="1" tabindex="5" accesskey="L" class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="5" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="6" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="7" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=rsServiceType.Fields.Item("insService_type_id").Value %>">
</form>
</body>
</html>
<%
rsServiceType.Close();
%>