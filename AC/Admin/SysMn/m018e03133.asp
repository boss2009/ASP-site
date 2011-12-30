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
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var rsRelationshipContactType = Server.CreateObject("ADODB.Recordset");
	rsRelationshipContactType.ActiveConnection = MM_cnnASP02_STRING;
	rsRelationshipContactType.Source = "{call dbo.cp_relationship_contact_type("+Request.QueryString("insRelationship_Contact_Type_Id")+",'"+ Description + "'," + IsActive + ",0,'E',0)}";
	rsRelationshipContactType.CursorType = 0;
	rsRelationshipContactType.CursorLocation = 2;
	rsRelationshipContactType.LockType = 3;
	rsRelationshipContactType.Open();
	Response.Redirect("m018q03133.asp");	
}

var rsRelationshipContactType = Server.CreateObject("ADODB.Recordset");
rsRelationshipContactType.ActiveConnection = MM_cnnASP02_STRING;
rsRelationshipContactType.Source = "{call dbo.cp_relationship_contact_type("+ Request.QueryString("insRelationship_Contact_Type_Id") + ",'',0,1,'Q',0)}";
rsRelationshipContactType.CursorType = 0;
rsRelationshipContactType.CursorLocation = 2;
rsRelationshipContactType.LockType = 3;
rsRelationshipContactType.Open();
%>
<html>
<head>
	<title>Update Contact Relationship Lookup</title>
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
				document.frm03133.reset();
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
<form name="frm03133" method="POST" action="<%=MM_editAction%>">
<h5>Update Contact Relationship Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Description:</td>
		<td nowrap><input type="text" name="Description" value="<%=(rsRelationshipContactType.Fields.Item("chvRelationship_Contact_Type").Value)%>" maxlength="40" size="40" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td nowrap>Is Active:</td>
		<td nowrap><input type="checkbox" name="IsActive" <%=((rsRelationshipContactType.Fields.Item("bitIs_Active").Value == 1)?"CHECKED":"")%> value="1" tabindex="2" accesskey="L" class="chkstyle"></td>
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
<input type="hidden" name="MM_recordId" value="<%=rsRelationshipContactType.Fields.Item("insRelationship_Contact_Type_Id").Value%>">
</form>
</body>
</html>
<%
rsRelationshipContactType.Close();
%>