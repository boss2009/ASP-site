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
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var rsRelationshipContactType = Server.CreateObject("ADODB.Recordset");
	rsRelationshipContactType.ActiveConnection = MM_cnnASP02_STRING;
	rsRelationshipContactType.Source = "{call dbo.cp_relationship_contact_type(0,'"+ Description + "'," + IsActive + ",0,'A',0)}";
	rsRelationshipContactType.CursorType = 0;
	rsRelationshipContactType.CursorLocation = 2;
	rsRelationshipContactType.LockType = 3;
	rsRelationshipContactType.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}
%>
<html>
<head>
	<title>New Contact Relationship Type</title>
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
		if (Trim(document.frm03133.Description.value)==""){
			alert("Enter Description.");
			document.frm03133.Description.focus();
			return ;		
		}
		document.frm03133.submit();
	}
	</script>	
</head>
<body onLoad="document.frm03133.Description.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm03133">
<h5>New Contact Relationship Type</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Description:</td>
		<td><input type="text" name="Description" maxlength="50" size="30" tabindex="1" accesskey="F"></td>
    </tr>
    <tr> 
		<td>Is Active:</td> 
		<td><input type="checkbox" name="IsActive" value="1" tabindex="2" accesskey="L" class="chkstyle"></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="4" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>